#!/usr/bin/env python3
"""MIG-01.5: read-only pallet/processing-chain scope audit of a legacy XLSM.

All reports are privacy-safe: ONLY table/ordinal, aggregate counts, enum labels,
column names and flags. No raw names, IDs, dates, business numbers, quantities,
amounts or account details. Nothing is written to the source. The classifications
are diagnostic hypotheses; this tool NEVER authorizes deletions/migration.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, Table, is_active, parse_xlsm, sha256_file, trim
from legacy_production_scope import KEPT, REVIEW, TEST, UNKNOWN, parse_scope
from legacy_test_scope import classify_tables

DOMAIN = ('tblPaleta','tblPaletaStavka','tblPrerada','tblPreradaStavka')
HEADERS = ('tblPaleta', 'tblPrerada')
STRONG = {'repo_test_seed','structured_test_token'}
PK = {
    'tblPaleta': 'PaletaID', 'tblPaletaStavka': 'StavkaID',
    'tblPrerada': 'PreradaID', 'tblPreradaStavka': 'StavkaID',
    'tblPrijemnica': 'PrijemnicaID', 'tblZbirna':'ZbirnaID',
}


def byname(tables:dict[str,Table], name:str) -> Table|None:
    return tables.get(name.lower())


def pk_index(table:Table|None, col:str)->dict[str,list[int]]:
    result=defaultdict(list)
    if not table or col not in table.columns: return result
    for ordinal,row in enumerate(table.rows,1):
        key=trim(row.get(col))
        if key:result[key].append(ordinal)
    return result


def classify_ref(key:str, index:dict[str,list[int]], scope:list[str]|None)->str:
    if not key:return 'NO_ID'
    hits=index.get(key, [])
    if not hits:return 'NO_TARGET'
    if len(hits)>1:return 'DUPLICATE_TARGET_ID'
    if scope is None:return 'FOUND_SCOPE_UNKNOWN'
    role=scope[hits[0]-1]
    return 'FOUND_'+({KEPT:'CORE', TEST:'TEST', REVIEW:'REVIEW',UNKNOWN:'UNCLASSIFIED'}.get(role,role))


def receipt_index_by_number(t:Table|None)->dict[str,list[int]]:
    d=defaultdict(list)
    if not t:return d
    for i,r in enumerate(t.rows,1):
        k=trim(r.get('BrojPrijemnice'))
        if k:d[k].append(i)
    return d


def business_no_hint(row:dict[str,str], bynumber:dict[str,list[int]], scope:list[str]|None)->str:
    """Business number is *not* authoritative for relinking; hint only."""
    k=trim(row.get('BrojPrijemnice'))
    if not k:return 'NO_BUSINESS_NUMBER'
    hits=bynumber.get(k,[])
    if not hits:return 'NOT_FOUND_BY_NUMBER'
    if scope is None:return 'NUMBER_FOUND_SCOPE_UNKNOWN'
    candidates=[i for i in hits if scope[i-1]==KEPT]
    if not candidates:return 'NO_CORE_CANDIDATE_BY_NUMBER'
    if len(candidates)>1:return 'MULTIPLE_CORE_ROWS_BY_NUMBER'
    return 'SINGLE_CORE_ROW_BY_NUMBER_NOT_VERIFIED'


def limit_examples(records:list[dict], max_count:int=80)->dict:
    return {'examples':records[:max_count], 'examples_truncated':max(0,len(records)-max_count)}


def evaluate(tables:dict[str,Table], proposal:dict, sha:str)->dict:
    scopes=parse_scope(proposal,tables,sha)
    _, signals=classify_tables(tables)
    for n in DOMAIN:
        if byname(tables,n) is None:
            raise AuditError('Nedostaje potrebna tabela: '+n)
    receipt=byname(tables,'tblPrijemnica')
    receipt_scope=scopes.get('tblPrijemnica')
    receipt_by_id=pk_index(receipt,'PrijemnicaID')
    receipt_by_no=receipt_index_by_number(receipt)
    paleta=byname(tables,'tblPaleta')
    paleta_stavka=byname(tables,'tblPaletaStavka')
    prerada=byname(tables,'tblPrerada')
    prerada_stavka=byname(tables,'tblPreradaStavka')
    ix_pal=pk_index(paleta,'PaletaID')
    ix_pre=pk_index(prerada,'PreradaID')

    result={
       'format':'MIG01.5-COLDCHAIN-READONLY',
       'privacy':'NO_RAW_IDS_NAMES_NUMBERS_DATES_MONEY_WEIGHTS_OR_BANK_VALUES',
       'source_sha256':sha,
       'migration_authorized':False,
       'scope_decision':'Paleta + PaletaStavka + Prerada + PreradaStavka ARE MANDATORY migration domains, never blanket OUT_OF_SCOPE',
       'receipt_scope':'operator candidate rows only; ID links to receipt do not prove pallet/processing business provenance',
       'tables':{},'relationships':{},'pallet_candidates':{},'processing_candidates':{},
       'storno_evidence':{},'issues':{},
    }
    for n in DOMAIN:
        t=byname(tables,n)
        signal_rows=signals.get(n,{})
        strong=[i for i,c in signal_rows.items() if STRONG & c]
        weak=[i for i,c in signal_rows.items() if 'descriptive_test_word' in c]
        result['tables'][n]={
          'rows':len(t.rows),'active':sum(is_active(row) for row in t.rows),
          'stornirane':sum(not is_active(row) for row in t.rows),
          'strong_test_markers':len(strong),'strong_marker_ordinals':strong[:30],
          'weak_test_markers':len(weak),'weak_marker_ordinals':weak[:20],
          'columns_present':{c:c in t.columns for c in (
              ('PaletaID','PreradaID','PrijemnicaID','BrojPrijemnice','Stornirano'))},
          'operator_scope_status':'NOT_PROVIDED_REVIEW_ALL',
        }

    # Pallet items: explicit parent PaletaID, direct receipt PrijemnicaID;
    # fall back to business numbers as non-authorizing diagnostics only.
    rel_counter=Counter(); problematic=[]; children_by_pal=defaultdict(list); known_real_receipt_by_pal=defaultdict(set)
    for i,r in enumerate(paleta_stavka.rows,1):
        pal_key=trim(r.get('PaletaID'))
        rec_key=trim(r.get('PrijemnicaID'))
        p_kind=classify_ref(pal_key,ix_pal,None)
        rec_kind=classify_ref(rec_key,receipt_by_id,receipt_scope)
        business=business_no_hint(r,receipt_by_no,receipt_scope)
        rel_counter[('parent_paleta',p_kind)]+=1
        rel_counter[('receipt_exact_id',rec_kind)]+=1
        if not rec_key:rel_counter[('receipt_business_number_non_authorizing',business)]+=1
        if pal_key in ix_pal and len(ix_pal[pal_key])==1:
            pal_idx=ix_pal[pal_key][0]
            children_by_pal[pal_idx].append((i,rec_kind,is_active(r),business))
            if rec_kind=='FOUND_CORE':known_real_receipt_by_pal[pal_idx].add(i)
        if p_kind not in ('FOUND_SCOPE_UNKNOWN',) or rec_kind not in ('FOUND_CORE',):
            problematic.append({'table':'tblPaletaStavka','row':i,'active':is_active(r),
                                'paleta_id_link':p_kind,'receipt_id_link':rec_kind,
                                'receipt_number_hint':business if rec_kind=='NO_ID' else 'NOT_USED'})
    result['relationships']['tblPaletaStavka']={
        'breakdown':[{'relation':a,'outcome':b,'count':v} for (a,b),v in sorted(rel_counter.items())],
        **limit_examples(problematic),
    }

    # Pallet classification does NOT decide KEEP/TEST. The old data may contain
    # experimental cancellation of otherwise genuine receipt-linked palettes.
    pc=Counter(); pallet_rows=[]
    for i,r in enumerate(paleta.rows,1):
        kids=children_by_pal.get(i,[])
        kflags={s for _j,s,_a,_b in kids}
        marked=bool(STRONG & signals.get('tblPaleta',{}).get(i,set()))
        if marked:bucket='STRONG_TEST_MARKER_REVIEW'
        elif 'FOUND_CORE' in kflags and any(s.startswith('FOUND_TEST') for s in kflags):bucket='MIXED_PROD_TEST_RECEIPTS_REVIEW'
        elif 'FOUND_CORE' in kflags:bucket='LINKED_TO_KEPT_RECEIPT_CANDIDATE'
        elif any(s in ('NO_ID','NO_TARGET','DUPLICATE_TARGET_ID') for s in kflags):bucket='UNRESOLVED_RECEIPT_IDENTITY_REVIEW'
        elif kids:bucket='LINKED_ONLY_TO_NONCORE_OR_UNKNOWN_REVIEW'
        else:bucket='NO_ITEMS_REVIEW'
        pc[(bucket,'ACTIVE' if is_active(r) else 'STORNO')]+=1
        if bucket!='LINKED_TO_KEPT_RECEIPT_CANDIDATE' and len(pallet_rows)<120:
            pallet_rows.append({'row':i,'active':is_active(r),'candidate':bucket,
                                'child_rows':len(kids), 'active_child_rows':sum(1 for x in kids if x[2]),
                                'strong_test_marker':marked})
    result['pallet_candidates']={'by_category':[{'classification':b,'state':s,'count':v} for (b,s),v in sorted(pc.items())],
                                  'interpretation':'All categories require operator approval; receipt link is evidence, never proof of production provenance.',
                                  **limit_examples(pallet_rows)}

    # Processing items are anchored by PreradaID -> Prerada and PaletaID -> Paleta.
    proc_counter=Counter(); proc_problem=[]; kids_by_pre=defaultdict(list)
    for i,r in enumerate(prerada_stavka.rows,1):
        pre_key=trim(r.get('PreradaID'))
        pal_key=trim(r.get('PaletaID'))
        pre_kind=classify_ref(pre_key,ix_pre,None)
        pal_kind=classify_ref(pal_key,ix_pal,None)
        proc_counter[('parent_prerada',pre_kind)]+=1
        proc_counter[('source_paleta',pal_kind)]+=1
        if pre_key in ix_pre and len(ix_pre[pre_key])==1:
            pr_i=ix_pre[pre_key][0]
            pal_i=ix_pal[pal_key][0] if pal_key in ix_pal and len(ix_pal[pal_key])==1 else None
            kids_by_pre[pr_i].append((i,pal_i,is_active(r)))
        if pre_kind!='FOUND_SCOPE_UNKNOWN' or pal_kind!='FOUND_SCOPE_UNKNOWN':
            proc_problem.append({'table':'tblPreradaStavka','row':i,'active':is_active(r),
                                 'prerada_id_link':pre_kind,'paleta_id_link':pal_kind})
    result['relationships']['tblPreradaStavka']={'breakdown':[{'relation':a,'outcome':b,'count':v} for (a,b),v in sorted(proc_counter.items())],
                                                 **limit_examples(proc_problem)}

    prc=Counter(); pre_details=[]
    for i,r in enumerate(prerada.rows,1):
        kids=kids_by_pre.get(i,[])
        pal_rows={x[1] for x in kids if x[1] is not None}
        real_linked_pal={p for p in pal_rows if p in known_real_receipt_by_pal}
        no_pal=sum(x[1] is None for x in kids)
        marked=bool(STRONG & signals.get('tblPrerada',{}).get(i,set()))
        if marked:kind='STRONG_TEST_MARKER_REVIEW'
        elif no_pal:kind='MISSING_PALETTE_LINK_REVIEW'
        elif not kids:kind='NO_ITEMS_REVIEW'
        elif real_linked_pal and len(real_linked_pal)==len(pal_rows):kind='ALL_SOURCE_PALETTES_LINKED_TO_KEPT_RECEIPTS_CANDIDATE'
        elif real_linked_pal:kind='MIXED_PROVENANCE_PALETTES_REVIEW'
        else:kind='UNPROVEN_SOURCE_PALETTE_PROVENANCE_REVIEW'
        prc[(kind,'ACTIVE' if is_active(r) else 'STORNO')]+=1
        if kind!='ALL_SOURCE_PALETTES_LINKED_TO_KEPT_RECEIPTS_CANDIDATE' and len(pre_details)<120:
            pre_details.append({'row':i,'active':is_active(r),'candidate':kind,'source_palettes':len(pal_rows),
                                'active_source_items':sum(bool(x[2]) for x in kids),
                                'strong_test_marker':marked})
    result['processing_candidates']={'by_category':[{'classification':b,'state':s,'count':v} for (b,s),v in sorted(prc.items())],
                                      'interpretation':'Stornirana prerada can be real history; do not exclude merely for storno status or position.',
                                      **limit_examples(pre_details)}

    # Evidence of storno events against palette/processing IDs. These are cues,
    # not evidence that the corresponding rows are tests. Do not copy history
    # blindly: some records mutate associated pallet/processing states.
    j=byname(tables,'tblStornoZurnal')
    zv=byname(tables,'tblStornoVeze')
    journal_matches=Counter()
    if j:
        for row in j.rows:
            table=trim(row.get('Tabela')).lower()
            if table in {x.lower() for x in DOMAIN}:
                journal_matches[next((x for x in DOMAIN if x.lower()==table),'UNKNOWN')]+=1
    exact_links=Counter()
    if zv:
        ix={n:pk_index(byname(tables,n),PK[n]) for n in HEADERS}
        for row in zv.rows:
            for typ,id_col in (('Old','OldDocID'),('New','NewDocID')):
                v=trim(row.get(id_col))
                if not v:continue
                for n in HEADERS:
                    if v in ix[n]:exact_links[(n,typ)]+=1
    result['storno_evidence']={'journal_rows_by_domain':dict(journal_matches),
                               'correction_context_id_links':[{'table':t,'role':role,'count':v}
                                                               for (t,role),v in sorted(exact_links.items())],
                               'note':'A storno journal may be incomplete. Confirm current row state and linked child rows; do not use journal alone to choose exclusions.'}
    # Storno lineage cross-state: an active child of a storniran header is not
    # automatically invalid. It is a forensic signal requiring operator review.
    consistency=Counter(); examples=[]
    for i,r in enumerate(paleta_stavka.rows,1):
        if not is_active(r):continue
        palid=trim(r.get('PaletaID'))
        recid=trim(r.get('PrijemnicaID'))
        ph=ix_pal.get(palid,[])
        rh=receipt_by_id.get(recid,[])
        if len(ph)==1 and not is_active(paleta.rows[ph[0]-1]):
            consistency['ACTIVE_PALETTE_ITEM_UNDER_STORNIRANA_PALETA']+=1
            examples.append({'table':'tblPaletaStavka','row':i,'case':'ACTIVE_PALETTE_ITEM_UNDER_STORNIRANA_PALETA','related_row':ph[0]})
        if len(rh)==1 and not is_active(receipt.rows[rh[0]-1]):
            consistency['ACTIVE_PALETTE_ITEM_LINKED_TO_STORNIRANA_PRIJEMNICA']+=1
            examples.append({'table':'tblPaletaStavka','row':i,'case':'ACTIVE_PALETTE_ITEM_LINKED_TO_STORNIRANA_PRIJEMNICA','related_row':rh[0]})
    for i,r in enumerate(prerada_stavka.rows,1):
        if not is_active(r):continue
        ph=ix_pal.get(trim(r.get('PaletaID')),[])
        prh=ix_pre.get(trim(r.get('PreradaID')),[])
        if len(ph)==1 and not is_active(paleta.rows[ph[0]-1]):
            consistency['ACTIVE_PROCESSING_ITEM_FROM_STORNIRANA_PALETA']+=1
            examples.append({'table':'tblPreradaStavka','row':i,'case':'ACTIVE_PROCESSING_ITEM_FROM_STORNIRANA_PALETA','related_row':ph[0]})
        if len(prh)==1 and not is_active(prerada.rows[prh[0]-1]):
            consistency['ACTIVE_PROCESSING_ITEM_UNDER_STORNIRANA_PRERADA']+=1
            examples.append({'table':'tblPreradaStavka','row':i,'case':'ACTIVE_PROCESSING_ITEM_UNDER_STORNIRANA_PRERADA','related_row':prh[0]})
    for i,r in enumerate(prerada.rows,1):
        if is_active(r) and not any(x[2] for x in kids_by_pre.get(i,[])):
            consistency['ACTIVE_PRERADA_WITHOUT_ACTIVE_ITEMS']+=1
            examples.append({'table':'tblPrerada','row':i,'case':'ACTIVE_PRERADA_WITHOUT_ACTIVE_ITEMS'})
    for i,r in enumerate(paleta.rows,1):
        if is_active(r) and not any(x[2] for x in children_by_pal.get(i,[])):
            consistency['ACTIVE_PALETA_WITHOUT_ACTIVE_ITEMS']+=1
            examples.append({'table':'tblPaleta','row':i,'case':'ACTIVE_PALETA_WITHOUT_ACTIVE_ITEMS'})
    result['cross_state_review']={
        'counts_by_case':dict(sorted(consistency.items())),
        **limit_examples(examples, 120),
        'interpretation':'Possible historical/cancellation states, never proof of defective workbook or tests.',
    }

    result['issues']={
        'cannot_auto_migrate_pallets':True,
        'cannot_auto_migrate_processing':True,
        'needs_operator_identification_of_real_pallet_and_processing_ids':True,
        'source_business_scope_kept_through_receipt':len(receipt.rows) if receipt else None,
        'unclassified_pallet_rows':len(paleta.rows),
        'unclassified_pallet_items':len(paleta_stavka.rows),
        'unclassified_processing_rows':len(prerada.rows),
        'unclassified_processing_items':len(prerada_stavka.rows),
    }
    result['limitations']=[
        'The operator has not yet supplied verified real/test boundaries for palette and processing tables.',
        'A link to an in-scope receipt does not prove a palette was a business operation: a test can operate on real receipts.',
        'A storno may be real business history; STORNO != TEST.',
        'Receipt number lookup is a hint only. ID+ownership+generation must be resolved before conversion.',
        'Migration would need to map legacy receipt IDs to new canonical IDs and validate active pallet/processing aggregates.',
        'Historical cancellation/reprocessing semantic validation is not implemented in MIG-01.5.',
    ]
    return result


def human(d:dict)->str:
    lines=['AGRIX MIG-01.5 - PALETA & PRERADA / STORNO FORENSIC AUDIT', '='*70,
           'READ ONLY. All 4 tables mandatory; NO automatic exclusion or migration.',
           'No raw IDs, names, dates, business numbers, weights, or financial data.', '']
    for table,summary in d['tables'].items():
        lines.append(f"{table:19} all={summary['rows']:5} active={summary['active']:5} "
                     f"storno={summary['stornirane']:5} strong_test_markers={summary['strong_test_markers']:3} "
                     f"weak_test_markers={summary['weak_test_markers']:3}")
    for table,rel in d['relationships'].items():
        lines.append('\n'+table+' -- typed identity integrity')
        for e in rel['breakdown']:
            lines.append(f"  {e['relation']}: {e['outcome']} = {e['count']}")
        lines.append('  flagged item ordinals (first %d): '%len(rel['examples'])+
                     ', '.join(str(e['row']) for e in rel['examples']))
    for key in ('pallet_candidates','processing_candidates'):
        lines.append('\n'+key.upper()+' -- triage only, never approval')
        for item in d[key]['by_category']:
            lines.append(f"  {item['classification']} [{item['state']}]: {item['count']}")
        lines.append('  review examples (first 60): '+', '.join(str(x['row']) for x in d[key]['examples'][:60]))
    lines.append('\nCROSS-STATE REVIEW (NOT AUTOMATIC ERRORS)')
    for k,v in d['cross_state_review']['counts_by_case'].items():
        lines.append(f'  {k}: {v}')
    lines.append('\nSTORNO JOURNAL matched domain records: '+str(d['storno_evidence']['journal_rows_by_domain']))
    for e in d['storno_evidence']['correction_context_id_links']:
        lines.append(f"  Correction: {e['table']} {e['role']} = {e['count']}")
    lines.append('\nDECISION: DO NOT ERASE STORNO ROWS, DO NOT DROP PALETA/PRERADA, '
                 'DO NOT USE LAST-ROW CUTS WITHOUT GRAPH VALIDATION.')
    lines.append('NEXT: operator validates actual historical operations, not mere test marker or source chronology.')
    return '\n'.join(lines)+'\n'


def main(argv=None):
    ap=argparse.ArgumentParser(description='MIG-01.5 read-only palette/processing forensic audit')
    ap.add_argument('--source',required=True)
    ap.add_argument('--out',required=True)
    ap.add_argument('--scope',default=str(Path(__file__).with_name('agrix_scope_operator_proposal.json')))
    a=ap.parse_args(argv)
    src=Path(a.source).expanduser().resolve()
    dest=Path(a.out).expanduser().resolve()
    if not src.is_file() or src.suffix.lower()!='.xlsm':ap.error('Izvor mora biti postojeca .xlsm sveska')
    if dest==src or dest==src.parent:ap.error('Izlaz mora biti poseban folder')
    try:
        proposal=json.loads(Path(a.scope).read_text(encoding='utf-8'))
        sha=sha256_file(src)
        d=evaluate(parse_xlsm(src),proposal,sha)
        dest.mkdir(parents=True,exist_ok=True)
        (dest/'agrix_coldchain_diagnostics.json').write_text(json.dumps(d,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        summary=human(d)
        (dest/'agrix_coldchain_diagnostics.txt').write_text(summary,encoding='utf-8')
        print(summary)
        print('Saved in:',dest)
        return 0
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as exc:
        print('STOP: dijagnostika nije uspela:',type(exc).__name__,str(exc)[:250],file=sys.stderr)
        return 2

if __name__=='__main__':sys.exit(main())
