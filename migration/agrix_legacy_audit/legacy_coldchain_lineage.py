#!/usr/bin/env python3
"""MIG-01.6: read-only lineage for pallets and processing around canceled receipts.

No raw workbook values are exported. All suggestions are forensic evidence, NEVER
an authorization to relink, delete, migrate or mutate historical records.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, Table, is_active, parse_xlsm, sha256_file, trim
from legacy_production_scope import parse_scope, KEPT, TEST, REVIEW, UNKNOWN
from legacy_coldchain_scope import byname, pk_index
from legacy_test_scope import classify_tables

STRONG = {'repo_test_seed', 'structured_test_token'}
REQUIRED = ('tblPaleta', 'tblPaletaStavka', 'tblPrerada', 'tblPreradaStavka', 'tblPrijemnica')


def active_scope(value: str) -> str:
    return {KEPT:'CORE_CANDIDATE',TEST:'TEST_CANDIDATE',REVIEW:'REVIEW',UNKNOWN:'UNKNOWN'}.get(value,'UNKNOWN')


def compatible(old: dict[str,str], new:dict[str,str]) -> bool:
    """Basic guard only, not evidence of semantic equivalence."""
    for key in ('KupacID', 'Klasa', 'VrstaVoca', 'SortaVoca'):
        a, b = trim(old.get(key)), trim(new.get(key))
        if a and b and a.casefold()!=b.casefold():
            return False
    return True


def successor_evidence(idx: int, receipts:Table, by_id:dict[str,list[int]],
                       scope:list[str]) -> dict:
    """Only explicit identity references authorize an *evidence classification*.

    No inferred link by number; even exact-target match is not auto-relink approval.
    """
    old=receipts.rows[idx-1]
    old_id=trim(old.get('PrijemnicaID'))
    exact=set();dangling=False
    direct=trim(old.get('ZamenjenSa'))
    if direct:
        hits=by_id.get(direct,[])
        if len(hits)==1 and hits[0]!=idx: exact.add(hits[0])
        else:dangling=True
    if old_id:
        for j,new in enumerate(receipts.rows,1):
            if j!=idx and trim(new.get('IspravkaOd'))==old_id:
                exact.add(j)
    active=[j for j in exact if is_active(receipts.rows[j-1])]
    candidates=[j for j in active if scope[j-1]==KEPT and compatible(old,receipts.rows[j-1])]
    if len(candidates)==1 and len(active)==1:
        result='UNIQUE_EXPLICIT_ACTIVE_CORE_SUCCESSOR_CANDIDATE'
    elif len(active)>1:
        result='MULTIPLE_EXPLICIT_ACTIVE_SUCCESSORS_REVIEW'
    elif active and not candidates:
        result='EXPLICIT_SUCCESSOR_OUT_OF_SCOPE_OR_MISMATCH'
    elif exact:
        result='EXPLICIT_SUCCESSOR_ONLY_STORNIRANO'
    elif dangling:
        result='BROKEN_EXPLICIT_SUCCESSOR_POINTER'
    elif trim(old.get('CorrectionID')):
        result='CORRECTION_ID_ONLY_NO_EXPLICIT_SUCCESSOR'
    else:
        result='NO_EXPLICIT_SUCCESSOR'
    return {'status':result, 'target_rows':sorted(exact)[:10],
            'target_count':len(exact),'active_target_count':len(active),
            'candidate_row':candidates[0] if result=='UNIQUE_EXPLICIT_ACTIVE_CORE_SUCCESSOR_CANDIDATE' else None,
            'broken_direct_pointer':dangling}


def analyze(tables:dict[str,Table], scope_cfg:dict, source_sha:str) -> dict:
    scopes=parse_scope(scope_cfg,tables,source_sha)
    _, markers=classify_tables(tables)
    for name in REQUIRED:
        if not byname(tables,name):raise AuditError('Nedostaje obavezna tabela: '+name)
    receipt=byname(tables,'tblPrijemnica')
    pallet=byname(tables,'tblPaleta')
    pallet_item=byname(tables,'tblPaletaStavka')
    proc=byname(tables,'tblPrerada')
    proc_item=byname(tables,'tblPreradaStavka')
    pr_scope=scopes['tblPrijemnica']
    ir=pk_index(receipt,'PrijemnicaID')
    ip=pk_index(pallet,'PaletaID')
    ix=pk_index(proc,'PreradaID')
    pal_by_receipt=defaultdict(list)
    pal_by_proc=defaultdict(set)
    proc_by_pal=defaultdict(set)
    proc_items_by_parent=defaultdict(list)
    receipt_by_pallet=defaultdict(list)
    for j,r in enumerate(pallet_item.rows,1):
        rh=ir.get(trim(r.get('PrijemnicaID')),[])
        ph=ip.get(trim(r.get('PaletaID')),[])
        if len(rh)==1 and len(ph)==1:
            receipt_by_pallet[ph[0]].append((j,rh[0],is_active(r)))
            pal_by_receipt[rh[0]].append((j,ph[0],is_active(r)))
    for j,r in enumerate(proc_item.rows,1):
        ph=ip.get(trim(r.get('PaletaID')),[])
        dh=ix.get(trim(r.get('PreradaID')),[])
        if len(ph)==1 and len(dh)==1:
            proc_by_pal[ph[0]].add(dh[0])
            pal_by_proc[dh[0]].add(ph[0])
            proc_items_by_parent[dh[0]].append((j,ph[0],is_active(r)))

    bad=[]; success_counts=Counter(); scope_counts=Counter(); pal_affected=set(); rec_affected=set(); proc_affected=set(); marker_counts=Counter()
    for i,r in enumerate(pallet_item.rows,1):
        if not is_active(r):continue
        rh=ir.get(trim(r.get('PrijemnicaID')),[])
        ph=ip.get(trim(r.get('PaletaID')),[])
        if len(rh)!=1 or is_active(receipt.rows[rh[0]-1]):continue
        rec=receipt.rows[rh[0]-1]; origin_scope=active_scope(pr_scope[rh[0]-1])
        ev=successor_evidence(rh[0],receipt,ir,pr_scope)
        pal_row=ph[0] if len(ph)==1 else None
        family=receipt_by_pallet.get(pal_row,[]) if pal_row else []
        active_test_items=sum(active and pr_scope[ri-1]==TEST for _j,ri,active in family)
        any_test_items=sum(pr_scope[ri-1]==TEST for _j,ri,active in family)
        pallet_marked=bool(pal_row and STRONG & markers.get('tblPaleta',{}).get(pal_row,set()))
        item_marked=bool(STRONG & markers.get('tblPaletaStavka',{}).get(i,set()))
        used_proc=proc_by_pal.get(pal_row,set()) if pal_row else set()
        active_proc=sum(is_active(proc.rows[j-1]) for j in used_proc)
        bad.append({'item_row':i,'receipt_row':rh[0], 'receipt_scope':origin_scope,
                    'pallet_row':pal_row, 'pallet_active':is_active(pallet.rows[pal_row-1]) if pal_row else None,
                    'pallet_strong_test_marker':pallet_marked,'item_strong_test_marker':item_marked,
                    'test_receipt_items_on_same_pallet':any_test_items,
                    'active_test_receipt_items_on_same_pallet':active_test_items,
                    'linked_processing_headers':len(used_proc),'active_processing_headers':active_proc,
                    'successor_evidence':ev['status'], 'explicit_successor_candidate_row':ev['candidate_row']})
        success_counts[ev['status']]+=1
        scope_counts[origin_scope]+=1
        if pal_row:pal_affected.add(pal_row)
        rec_affected.add(rh[0]);proc_affected.update(used_proc)
        if pallet_marked:marker_counts['pallet_strong_marker']+=1
        if item_marked:marker_counts['item_strong_marker']+=1
        if any_test_items:marker_counts['has_test_receipt_items']+=1

    # Every processing header is examined; do not mistake 'linked to any kept
    # receipt' for clean provenance when same pallet also carries TEST receipts.
    processing=Counter(); processing_details=[]
    for i,r in enumerate(proc.rows,1):
        item_links=proc_items_by_parent.get(i,[])
        parents={p for _j,p,_a in item_links}
        strong_pals=[p for p in parents if STRONG & markers.get('tblPaleta',{}).get(p,set())]
        mixed=[]; untrusted=[]; past_storn_receipt=[]
        for p in sorted(parents):
            items=receipt_by_pallet.get(p,[])
            statuses={pr_scope[rec-1] for _j,rec,_active in items}
            if TEST in statuses and KEPT in statuses:mixed.append(p)
            if not items or (KEPT not in statuses):untrusted.append(p)
            if any(active and not is_active(receipt.rows[rec-1]) for _j,rec,active in items):
                past_storn_receipt.append(p)
        issues=[]
        if strong_pals:issues.append('SOURCE_PALLET_TEST_MARKER')
        if mixed:issues.append('SOURCE_PALLET_MIXES_CORE_TEST')
        if untrusted:issues.append('SOURCE_PALLET_WITHOUT_CORE_RECEIPT')
        if past_storn_receipt:issues.append('SOURCE_PALLET_ACTIVE_ITEM_ON_STORNIRANA_RECEIPT')
        if not parents:issues.append('NO_VALID_LINKED_PALLET_ITEMS')
        if i in proc_affected:issues.append('TOUCHES_41_CASES')
        if not issues:issues=['NO_CURRENT_GRAPH_RED_FLAGS_NOT_VALIDATED']
        for issue in issues:processing[(issue,'ACTIVE' if is_active(r) else 'STORNO')]+=1
        if issues!=['NO_CURRENT_GRAPH_RED_FLAGS_NOT_VALIDATED']:
            processing_details.append({'processing_row':i,'active':is_active(r),
                'source_pallet_count':len(parents),'flags':issues,
                'pallet_row_ordinals':sorted(parents)[:20]})
    result={'format':'MIG01.6-COLDCHAIN-LINEAGE-READONLY','source_sha256':source_sha,
            'privacy':'ONLY_ORDINALS_COUNTS_FLAGS_NO_RAW_SOURCE_VALUES',
            'migration_authorized':False,
            'columns_available':{k:k in receipt.columns for k in ('PrijemnicaID','IspravkaOd','ZamenjenSa','CorrectionID','GeneracijaID','KupacID','Klasa','VrstaVoca','SortaVoca')},
            'active_pallet_items_to_stornirana_receipts':{
                 'count':len(bad),'distinct_receipts':len(rec_affected),
                 'distinct_pallets':len(pal_affected),'distinct_processing_headers':len(proc_affected),
                 'by_receipt_scope':dict(sorted(scope_counts.items())),
                 'by_successor_evidence':dict(sorted(success_counts.items())),
                 'by_additional_risk_flag':dict(sorted(marker_counts.items())),
                 'cases':bad,
            },
            'processing_graph_review':{
                'total_processing_headers':len(proc.rows),
                'flag_counts':[{'flag':k,'state':state,'count':n} for (k,state),n in sorted(processing.items())],
                'flagged_headers':processing_details,
                'caveat':'Counts of flags overlap; a processing header may have several flags.'
            },
            'interpretation':[
                'An explicit active successor is only a proposed identity candidate; never auto-reassign pallet items.',
                'No explicit successor does NOT prove a corrupt business event; storno may be intentional.',
                'Do not treat 97 processing rows with some core-receipt ancestry as proven production.',
                'Operator decisions must be approved for whole pallet/processing dependency chains, not just trailing rows.',
            ]}
    return result


def human(result:dict)->str:
    d=result['active_pallet_items_to_stornirana_receipts'];p=result['processing_graph_review']
    lines=['AGRIX MIG-01.6 - PALETA/PRERADA HISTORICAL LINEAGE', '='*64,
           'READ ONLY. No migration, deletion, macro execution or source modifications.',
           'No raw IDs, names, business numbers, dates, amounts or weights.',
           '','RECEIPT LINK REVIEW',
           f"  active pallet items to stornirana receipt: {d['count']}",
           f"  distinct receipt rows: {d['distinct_receipts']}",
           f"  distinct pallet rows: {d['distinct_pallets']}",
           f"  distinct processing headers affected: {d['distinct_processing_headers']}"]
    for k,n in d['by_receipt_scope'].items():lines.append(f'  receipt scope {k}: {n}')
    lines.append('EXPLICIT SUCCESSOR EVIDENCE (NOT APPROVAL)')
    for k,n in d['by_successor_evidence'].items():lines.append(f'  {k}: {n}')
    lines.append('ADDITIONAL RISK SIGNALS (OVERLAPPING)')
    for k,n in d['by_additional_risk_flag'].items():lines.append(f'  {k}: {n}')
    lines.append('\nCASES: pallet item / receipt / pallet / successor / downstream processing')
    for c in d['cases'][:85]:
        lines.append(f"  item={c['item_row']} rec={c['receipt_row']} ({c['receipt_scope']}) "
            f"pal={c['pallet_row']} pallet_active={c['pallet_active']} "
            f"test_children={c['test_receipt_items_on_same_pallet']} "
            f"proc={c['linked_processing_headers']} active_proc={c['active_processing_headers']} "
            f"successor={c['successor_evidence']} candidate_row={c['explicit_successor_candidate_row']}")
    lines.append('\nPROCESSING GRAPH REVIEW (FLAG COUNTS OVERLAP)')
    lines.append(f"  total processing headers: {p['total_processing_headers']}")
    for x in p['flag_counts']:lines.append(f"  {x['flag']} [{x['state']}]: {x['count']}")
    lines.append('  flagged processing header ordinals: '+', '.join(str(x['processing_row']) for x in p['flagged_headers'][:90]))
    lines.append('\nDECISION: remains NO-GO for automated pallet/processing migration. Review by whole connected operation.')
    return '\n'.join(lines)+'\n'


def main(argv=None):
    ap=argparse.ArgumentParser(description='Read-only targeted lineage audit of historical pallet/processing operations.')
    ap.add_argument('--source',required=True)
    ap.add_argument('--out',required=True)
    ap.add_argument('--scope',default=str(Path(__file__).with_name('agrix_scope_operator_proposal.json')))
    a=ap.parse_args(argv)
    src=Path(a.source).expanduser().resolve();out=Path(a.out).expanduser().resolve()
    if not src.is_file() or src.suffix.lower()!='.xlsm':ap.error('--source must be an existing .xlsm')
    if src==out or out==src.parent:ap.error('--out must be a separate folder')
    try:
        sha=sha256_file(src)
        scope=json.loads(Path(a.scope).read_text(encoding='utf-8'))
        result=analyze(parse_xlsm(src),scope,sha)
        out.mkdir(parents=True,exist_ok=True)
        (out/'agrix_coldchain_lineage.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        txt=human(result)
        (out/'agrix_coldchain_lineage.txt').write_text(txt,encoding='utf-8')
        print(txt)
        print('WROTE agrix_coldchain_lineage.txt and .json in',out)
        return 0
    except (AuditError,ValueError,OSError,KeyError) as err:
        print('STOP: lineage analysis failed:',type(err).__name__,str(err)[:200],file=sys.stderr)
        return 2

if __name__=='__main__':sys.exit(main())
