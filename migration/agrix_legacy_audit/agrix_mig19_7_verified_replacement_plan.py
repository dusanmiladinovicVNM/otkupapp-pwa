#!/usr/bin/env python3
"""MIG-19.7: fail-closed source-grounded OTKUP replacement candidates and in-memory parity rehearsal.

NEVER changes source JSONL or macro-enabled workbook. Produces explicit one-to-one replacement
manifest only when source source document has two opposite station/cooperant legs and target has
exactly one existing canonical event under the same document/date/pack, with compatible movement
kind. A manifest is a proposal, NOT a safe-to-write authorization: AMB-10 document/invariant and
entire-event lineage review still required before any XLSM change.
"""
from __future__ import annotations
import argparse, json, sys
from collections import defaultdict, Counter
from decimal import Decimal
from pathlib import Path
from agrix_mig19_3_parity import (STAGE_SHA, XlsmReadOnly, stage_rows, ledger_views,
    reconcile, number, key, norm_type, date, void, dec, sha256, write_csv)
from agrix_mig19_4_difference_isolation import PRIOR_TARGET_SHA
from agrix_mig19_6_precise_manifest import doc_group, grouped_vectors, classify_doc, READ_COLUMNS, SEVEN
from agrix_mig19_malina_readonly_audit import MATRIX, account_class

OUTPUT_COLUMNS=['Status','Reason','Datum','TipAmbalaze','DokumentTIP','DokumentID',
   'SourceAmbIDs','TargetAmbID','ExistingVrstaKretanja','OldSource','OldDestination',
   'OldQty','TargetOldSource','TargetOldDestination','TargetOldQty',
   'ExpectedOdNalogTip','ExpectedOdNalogID','ExpectedNaNalogTip','ExpectedNaNalogID',
   'ExpectedQty','SourceRowCount','TargetNonCoverRowCount','TargetCoverRowCount']


def kind_valid(kind, src, dst):
    if kind not in MATRIX:return False
    want_a,want_b=MATRIX[kind]
    return want_a in account_class(src[0]) and want_b in account_class(dst[0])


def orig_pair(rows, doc):
    """Source evidence for one transfer: two separate ORIGINAL equal/opposite ledger legs."""
    if len(rows)!=2:return None,'LEGACY_NOT_TWO_LEGS'
    if {key(r.get('EntitetTip')) for r in rows}!={'STANICA','KOOPERANT'}:
        return None,'LEGACY_NOT_STATION_COOP_PAIR'
    amounts=[]
    for r in rows:
        amt=number(r.get('Kolicina'),'OTKUP source',positive=True)
        direction=key(r.get('Smer'))
        if direction not in ('ULAZ','IZLAZ'):return None,'SOURCE_INVALID_SMER'
        amounts.append(amt if direction=='ULAZ' else -amt)
    if amounts[0]+amounts[1]!=0 or amounts[0]==0:
        return None,'SOURCE_LEGS_NOT_BALANCED'
    neg=rows[0] if amounts[0]<0 else rows[1]
    pos=rows[0] if amounts[0]>0 else rows[1]
    src=(key(neg['EntitetTip']),key(neg['EntitetID']))
    dst=(key(pos['EntitetTip']),key(pos['EntitetID']))
    if not src[1] or not dst[1]:return None,'SOURCE_ACCOUNT_ID_EMPTY'
    return (src,dst,abs(amounts[0])),'OK'


def get_otkup_proposals(olds,news):
    oldg=defaultdict(list);newg=defaultdict(list);covered=Counter()
    for i,r in enumerate(olds,1):
        if void(r.get('Stornirano')):continue
        g=doc_group(r,i,'old#')
        if g[2]=='OTKUP':oldg[g].append(r)
    for i,r in enumerate(news,1):
        g=doc_group(r,i,'new#')
        if g[2]!='OTKUP':continue
        if key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE':covered[g]+=1
        else:newg[g].append(r)
    referenced_ids={key(r.get('StornoOd')) for r in news if key(r.get('StornoOd'))}
    ov,nv,_,_,_,_,_=grouped_vectors(olds,news)
    rows=[];candidates=[];status=Counter(); reasons=Counter(); all_docs=Counter()
    for g in sorted(set(oldg)|set(newg)):
        oldvec,newvec=ov.get(g,{}),nv.get(g,{})
        typ=classify_doc(oldvec,newvec)
        all_docs[typ]+=1
        if typ=='EXACT':continue
        source=oldg.get(g,[]);target=newg.get(g,[])
        record={f:'' for f in OUTPUT_COLUMNS}
        record.update(Datum=g[0],TipAmbalaze=g[1],DokumentTIP=g[2],DokumentID=g[3],
            SourceAmbIDs='|'.join(key(r.get('AmbID')) for r in source),
            SourceRowCount=len(source),TargetNonCoverRowCount=len(target),
            TargetCoverRowCount=covered[g])
        p,why=orig_pair(source,g)
        if why!='OK':reason=why
        elif len(target)!=1:reason='TARGET_NOT_SINGLE_PHYSICAL_ROW'
        elif covered[g]:reason='DERIVED_COVER_REQUIRES_INVARIANT_REVIEW'
        else:
            src,dst,q=p
            r=target[0]
            oldsrc=(key(r.get('OdNalogTip')),key(r.get('OdNalogID')))
            olddst=(key(r.get('NaNalogTip')),key(r.get('NaNalogID')))
            vk=key(r.get('VrstaKretanja'))
            record.update(TargetAmbID=key(r.get('AmbID')),ExistingVrstaKretanja=vk,
                OldSource=f'{src[0]}/{src[1]}',OldDestination=f'{dst[0]}/{dst[1]}',OldQty=dec(q),
                TargetOldSource=f'{oldsrc[0]}/{oldsrc[1]}',TargetOldDestination=f'{olddst[0]}/{olddst[1]}',
                TargetOldQty=dec(number(r.get('Kolicina'),'target OTKUP',positive=True)),
                ExpectedOdNalogTip=src[0],ExpectedOdNalogID=src[1],
                ExpectedNaNalogTip=dst[0],ExpectedNaNalogID=dst[1],ExpectedQty=dec(q))
            if set((src,dst))!=set((oldsrc,olddst)):
                reason='TARGET_PARTICIPANT_IDENTITY_DIFFERENT'
            elif key(r.get('StornoOd')) or key(r.get('AmbID')) in referenced_ids:
                reason='TARGET_HAS_STORNO_RELATION'
            elif not kind_valid(vk,src,dst):
                reason='EXISTING_MOVEMENT_KIND_INCOMPATIBLE_WITH_SOURCE_DIRECTION'
            elif not key(g[3]):
                reason='NO_DOCUMENT_ID'
            else:
                reason='ONE_TO_ONE_SOURCE_VECTOR_CANDIDATE'
                candidates.append((g,r,src,dst,q))
        record['Status']='CANDIDATE_NOT_APPROVED' if reason=='ONE_TO_ONE_SOURCE_VECTOR_CANDIDATE' else 'HOLD'
        record['Reason']=reason
        rows.append(record);reasons[reason]+=1;status[record['Status']]+=1
    return rows,candidates,reasons,all_docs


def rehearsal(olds,news,candidates):
    # Never write: simulate replacing exactly ONE original target row per candidate.
    patched=[dict(r) for r in news]
    id_index=defaultdict(list)
    for i,r in enumerate(patched):id_index[key(r.get('AmbID'))].append(i)
    for g,r,src,dst,q in candidates:
        aid=key(r.get('AmbID'))
        if len(id_index[aid])!=1:raise RuntimeError('Duplicate target AmbID: '+aid)
        entry=patched[id_index[aid][0]]
        assert doc_group(entry,id_index[aid][0]+1,'canonical#')==g
        entry.update(OdNalogTip=src[0],OdNalogID=src[1],NaNalogTip=dst[0],NaNalogID=dst[1],Kolicina=dec(q))
    return ledger_views(olds,patched)


def main(argv=None):
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--stage',required=True,type=Path)
    p.add_argument('--workbook',required=True,type=Path)
    p.add_argument('--out-dir',required=True,type=Path)
    args=p.parse_args(argv)
    stage=args.stage/'UNCONVERTED__tblAmbalaza.jsonl'
    for f in (stage,args.workbook):
        if not f.is_file():p.error('Missing required input '+str(f))
    out=args.out_dir.resolve();protected=[stage.resolve(),args.workbook.resolve()]
    if out in protected or out in [x.parent for x in protected]:p.error('Output must not be inside input parent')
    before={str(x):sha256(x) for x in protected}
    if before[str(stage.resolve())]!=STAGE_SHA:raise ValueError('Unexpected SOURCE SHA')
    if before[str(args.workbook.resolve())]!=PRIOR_TARGET_SHA:raise ValueError('Unexpected QA workbook SHA')
    old=stage_rows(stage);reader=XlsmReadOnly(args.workbook)
    try:new=reader.read_table('tblAmbalaza',READ_COLUMNS)
    finally:reader.close()
    orig=ledger_views(old,new)
    oacc,oday,odoc=reconcile(orig[0],orig[1],orig[2],orig[3],orig[4],orig[5])
    rows,can,reasons,types=get_otkup_proposals(old,new)
    test=rehearsal(old,new,can)
    tacc,tday,tdoc=reconcile(test[0],test[1],test[2],test[3],test[4],test[5])
    if before!={str(x):sha256(x) for x in protected}:raise RuntimeError('Source or QA workbook changed during audit')
    out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG197_ONE_TO_ONE_REPLACEMENT_CANDIDATES_PRIVATE.csv',
              [r for r in rows if r['Status']=='CANDIDATE_NOT_APPROVED'],OUTPUT_COLUMNS)
    write_csv(out/'MIG197_HELD_DOCUMENTS_PRIVATE.csv',
              [r for r in rows if r['Status']=='HOLD'],OUTPUT_COLUMNS)
    def tally(acc,dates,docs):
        return dict(account_type_mismatches=sum(r['Status']!='PASS' for r in acc),
                    cooperant_balance_mismatches=sum(r['Status']!='PASS' and r['NalogTip']=='KOOPERANT' for r in acc),
                    station_balance_mismatches=sum(r['Status']!='PASS' and r['NalogTip']=='STANICA' for r in acc),
                    daily_close_mismatches=sum(r['Status']!='PASS' for r in dates),
                    stable_document_account_slice_mismatches=sum(r['Status']!='PASS' for r in docs))
    summary={
        'format':'AGRIx-MIG19.7-VERIFIED-ONE-TO-ONE-REPLACEMENT-REHEARSAL-v1',
        'status':'CANDIDATE_MANIFEST_ONLY_NO_XLSM_WRITE',
        'inputs_unchanged':True,'stage_sha_verified':True,
        'qa_workbook_sha256':before[str(args.workbook.resolve())],
        'original_otkup_document_class_counts':dict(sorted(types.items())),
        'document_mismatch_reasons':dict(sorted(reasons.items())),
        'otkup_documents_different':len(rows),
        'one_to_one_replacement_candidates':len(can),
        'held_otkup_documents':len(rows)-len(can),
        'before':tally(oacc,oday,odoc),
        'in_memory_rehearsal_only':tally(tacc,tday,tdoc),
        'unresolved_seven_source_ids':sorted(SEVEN),
        'new_model_event_lineage_proven':False,
        'apply_authorized':False,
        'warning':'Only identical-source-document one-to-one candidates are previewed; full AMB10 invariants, other documents and cross-event mappings remain unproven. No workbook edited.'}
    (out/'MIG197_SUMMARY.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    return 0

if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print('MIG197_ERROR: '+type(exc).__name__+': '+str(exc),file=sys.stderr)
        sys.exit(3)
