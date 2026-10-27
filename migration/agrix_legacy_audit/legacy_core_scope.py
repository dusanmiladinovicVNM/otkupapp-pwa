#!/usr/bin/env python3
"""MIG-01.4: scope audit restricted to receiving/receipt chain + packaging + agro.

No migration, no row exports, no IDs, business numbers, dates or monetary amounts
in reports. Never modifies/opens the source workbook using Excel or macros.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET

from legacy_audit import AuditError, Table, is_active, parse_xlsm, sha256_file, trim
from legacy_diagnostics import group_key
from legacy_production_scope import (
    KEPT, REVIEW, TEST, UNKNOWN, FK_TO_TARGET, parse_scope,
    candidate_group_boundary,
)
from legacy_test_scope import PK_BY_TABLE, classify_tables

# Business-scope decisions from operator. Terminal endpoints (fakture/prerada/
# palete/utovar) are deliberately not dependencies of the requested migration.
CORE_CUTOFFS = {
    'tblOtkup', 'tblOtpremnica', 'tblZbirna', 'tblPrijemnica',
    'tblAmbalaza', 'tblMagacin', 'tblKooperanti', 'tblStanice',
    'tblKupci', 'tblVozaci', 'tblKulture', 'tblParcele',
}
CORE_MASTER_UNVERIFIED = {'tblArtikli', 'tblTipAmbalaze'}
SUPPORT_REVIEW = {'tblStornoVeze', 'tblStornoZurnal', 'tblCenovnik'}
FINANCE_SECONDARY = {'tblNovac', 'tblBankaImport', 'tblPartnerMap'}
OUT_OF_SCOPE = {
    'tblFakture', 'tblFakturaStavke', 'tblUtovar', 'tblUtovarStavke',
    'tblPrevoznici', 'tblPrerada', 'tblPreradaStavka',
    'tblPaleta', 'tblPaletaStavka', 'tblKese', 'tblKutije',
    'tblTipPalete', 'tblVrstaGotovihProizvoda',
}
REGENERATE = {'tblPoruke', 'tblConfig', 'tblSEFConfig',
              'tblLocalConfig', 'tblKorisnici', 'tblSEFEventLog',
              'tblSEFSubmission', 'tblMGMT'}
# Directional, typed references into the operational chain.
# The documented end-of-scope reference from a receipt to an invoice is ignored.
TERMINAL_ALLOWED = {('tblPrijemnica', 'FakturaID')}
KNOWN_TABLES = CORE_CUTOFFS | CORE_MASTER_UNVERIFIED | SUPPORT_REVIEW | FINANCE_SECONDARY | OUT_OF_SCOPE | REGENERATE


def tiers(tables: dict[str, Table], proposal: dict, source_sha: str) -> dict[str, list[str]]:
    original = parse_scope(proposal, tables, source_sha)
    result = {}
    for table in tables.values():
        name = table.name
        if name in CORE_CUTOFFS:
            if name not in proposal.get('last_real_row_inclusive', {}):
                raise AuditError('Nedostaje operatorova granica za CORE: ' + name)
            result[name] = [
                'CORE' if state == KEPT else
                'CORE_REVIEW' if state == REVIEW else
                'TEST_OUT' if state == TEST else
                'UNCLASSIFIED'
                for state in original[name]
            ]
        elif name in CORE_MASTER_UNVERIFIED:
            # Needed for foreign key integrity; absence of operator cutoff is
            # explicitly signaled rather than silently treated as approved.
            result[name] = ['MASTER_CANDIDATE'] * len(table.rows)
        elif name in SUPPORT_REVIEW:
            result[name] = ['SUPPORT_REVIEW'] * len(table.rows)
        elif name in FINANCE_SECONDARY:
            result[name] = [
                'FINANCE_CANDIDATE' if state == KEPT else
                'FINANCE_REVIEW' if state in (REVIEW, UNKNOWN) else
                'TEST_OUT'
                for state in original[name]
            ]
        elif name in OUT_OF_SCOPE:
            result[name] = ['OUT_OF_SCOPE'] * len(table.rows)
        elif name in REGENERATE:
            result[name] = ['REGENERATE_OR_RECONFIGURE'] * len(table.rows)
        else:
            result[name] = ['UNCLASSIFIED'] * len(table.rows)
    return result


def boundary_checks(tables: dict[str,Table], raw_state: dict[str,list[str]],
                    strong: dict[str,dict[int,set[str]]]) -> list[dict]:
    out=[]
    for crossing in candidate_group_boundary(tables,raw_state):
        name=crossing['table']
        if name not in CORE_CUTOFFS:continue
        table=tables[name.lower()]
        ordinal_list=crossing['ordinals']
        rows=[]
        for ix in ordinal_list:
            row=table.rows[ix-1]
            sign=strong.get(name,{}).get(ix,set())
            cls=trim(row.get('Klasa')).upper()
            rows.append({
                'ordinal':ix,
                'proposed_role':'CORE' if raw_state[name][ix-1]==KEPT else 'EXCLUDED',
                'active':is_active(row),
                'class':cls if cls in {'I','II'} else '(OTHER_OR_EMPTY)',
                'structured_test_hint':bool(sign & {'repo_test_seed','structured_test_token'}),
                'descriptive_test_hint':'descriptive_test_word' in sign,
                'has_generation_id':bool(trim(row.get('GeneracijaID'))),
            })
        out.append({'table':name,'basis':crossing['grouping_basis'],
                    'rows':rows,'decision':'NEEDS_OPERATOR_REVIEW',
                    'reason':('Explicit GeneracijaID crosses cutoff' if crossing['grouping_basis']=='GEN'
                              else 'Heuristic grouping crosses cutoff; not proof of same document')})
    return out


def core_crosslinks(tables: dict[str,Table], tier:dict[str,list[str]]) -> dict:
    index = {}
    for t in tables.values():
        pk=PK_BY_TABLE.get(t.name)
        if pk in t.columns:
            d=defaultdict(list)
            for i,row in enumerate(t.rows,1):
                v=trim(row.get(pk))
                if v:d[v].append(i)
            index[t.name]=d
    flags=Counter()
    examples=[]
    examined=0
    for table in tables.values():
        if table.name not in CORE_CUTOFFS and table.name not in CORE_MASTER_UNVERIFIED:
            continue
        pk=PK_BY_TABLE.get(table.name)
        for i,row in enumerate(table.rows,1):
            if tier[table.name][i-1] not in ('CORE','MASTER_CANDIDATE'):
                continue
            for col,target in FK_TO_TARGET.items():
                if col==pk or col not in row or not trim(row.get(col)):
                    continue
                if (table.name,col) in TERMINAL_ALLOWED:
                    continue
                if target not in index:
                    # If a master table is absent, this is a mandatory review.
                    flags[('TARGET_TABLE_NOT_INDEXED',table.name,col,target)] += 1
                    continue
                examined+=1
                tgt_rows=index[target].get(trim(row[col]),[])
                if not tgt_rows:
                    kind='FK_NOT_FOUND'
                elif any(tier[target][j-1] in ('CORE','MASTER_CANDIDATE') for j in tgt_rows):
                    continue
                else:
                    kind='TARGET_NOT_IN_CORE'
                flags[(kind,table.name,col,target)] += 1
                if len(examples)<70:
                    examples.append({'kind':kind,'source_table':table.name,
                                     'source_ordinal':i,'field':col,'target_table':target,
                                     'target_ordinals':tgt_rows[:8],
                                     'target_roles':[tier[target][j-1] for j in tgt_rows[:8]]})
    return {'validated_fk_cells':examined,
            'findings_total':sum(flags.values()),
            'by_type':[{'finding':k,'source_table':src,'field':col,'target_table':dst,'count':n}
                       for (k,src,col,dst),n in sorted(flags.items())],
            'examples':examples,
            'truncated_examples':max(0,sum(flags.values())-len(examples)),
            'note':'This checks exact typed FK values only. Business numbers, AMB account meaning and storno still require separate verification.'}


def unknown_zbirna(tables:dict[str,Table],tier:dict[str,list[str]])->list[dict]:
    z=tables.get('tblzbirna')
    if not z:return []
    groups=defaultdict(list)
    for j,r in enumerate(z.rows,1):
        number=trim(r.get('BrojZbirne'))
        if number:groups[number].append(j)
    anomalies=[]
    for table_name in ('tblOtkup','tblOtpremnica','tblPrijemnica'):
        t=tables.get(table_name.lower())
        if t is None:continue
        for i,row in enumerate(t.rows,1):
            if tier[table_name][i-1]!='CORE':continue
            number=trim(row.get('BrojZbirne'))
            if not number:continue
            hits=groups.get(number,[])
            kept=[j for j in hits if tier['tblZbirna'][j-1]=='CORE']
            if not hits: reason='NUMBER_NOT_FOUND'
            elif not kept: reason='ONLY_NON_CORE_CANDIDATES'
            else: continue
            anomalies.append({'source_table':table_name,'source_ordinal':i,
                              'active':is_active(row),'reason':reason,
                              'number_matches_total':len(hits),
                              'number_matches_kept':len(kept)})
    return anomalies


def evaluate(tables:dict[str,Table],scope:dict,sha:str)->dict:
    raw=parse_scope(scope,tables,sha)
    tier=tiers(tables,scope,sha)
    _, strong=classify_tables(tables)
    counts={n:dict(Counter(tier[n])) for n in sorted(tier) if any(tier[n])}
    splits=boundary_checks(tables,raw,strong)
    missing=unknown_zbirna(tables,tier)
    fks=core_crosslinks(tables,tier)
    terminal_counts = {}
    for table_name, field in sorted(TERMINAL_ALLOWED):
        table=tables.get(table_name.lower())
        if table:
            terminal_counts[table_name+'.'+field]=sum(
                tier[table_name][i-1]=='CORE' and bool(trim(row.get(field)))
                for i,row in enumerate(table.rows,1))
    all_names={t.name for t in tables.values()}
    unforeseen=sorted(all_names-KNOWN_TABLES)
    return {
        'format':'MIG01.4-CORE-SCOPE-READ-ONLY',
        'privacy':'NO_RAW_ROW_VALUES_IDS_NUMBERS_DATES_OR_AMOUNTS',
        'source_sha256':sha,
        'migration_authorized':False,
        'core_definition': 'Otkup -> Otpremnica -> Zbirna -> Prijemnica; ambalaza; agrohemija; required master records',
        'excluded_by_scope':sorted(OUT_OF_SCOPE & all_names),
        'regenerated_from_repo_or_reconfigured':sorted(REGENERATE & all_names),
        'secondary_not_blocking_core':sorted(FINANCE_SECONDARY & all_names),
        'support_requires_review':sorted(SUPPORT_REVIEW & all_names),
        'unknown_tables':unforeseen,
        'row_counts_by_role':counts,
        'boundary_documents':splits,
        'zbirna_number_not_found':missing,
        'core_typed_fk_evidence':fks,
        'terminal_references_to_omitted_domains':terminal_counts,
        'outstanding':{
            'cross_boundary_document_candidate_groups':len(splits),
            'unresolved_zbirna_number_references':len(missing),
            'core_fk_issues':fks['findings_total'],
            'agro_review_rows':counts.get('tblMagacin',{}).get('CORE_REVIEW',0),
            'core_master_rows_needing_scope_approval':sum(counts.get(n,{}).get('MASTER_CANDIDATE',0) for n in CORE_MASTER_UNVERIFIED),
            'storno_history_rows_needing_review':sum(len(tables[n.lower()].rows) for n in SUPPORT_REVIEW if n in ('tblStornoVeze','tblStornoZurnal') and n.lower() in tables),
        },
        'limitations':['No complete AMB-10 event/ownership mapping',
                       'No workbook row values are exported',
                       'Document grouping uses candidate GEN/HEUR, not proven identity',
                       'Financial values and status are not reconciled',
                       'Historical storno/correction chains still require inspection'],
    }


def human_report(result:dict)->str:
    lines=['AGRIX MIG-01.4 - CORE OPERATIONAL SCOPE', '='*55,
           'READ ONLY - NO MIGRATION - NO USER DATA EXPORTED',
           'Core: Otkup -> Otpremnica -> Zbirna -> Prijemnica; ambalaza; agrohemija.',
           'Fakture/palete/prerada/utovar are explicitly OUT OF SCOPE.',
           'Novac / BankaImport = SECONDARY, never auto-deleted or considered empty.', '']
    counts=result['row_counts_by_role']
    for name in sorted(CORE_CUTOFFS|CORE_MASTER_UNVERIFIED):
        if name in counts:
            lines.append(f"{name:22}: "+' '.join(f'{k}={v}' for k,v in sorted(counts[name].items())))
    lines.append('\nCROSS-BOUNDARY DOCUMENT CANDIDATES')
    for item in result['boundary_documents']:
        lines.append(f"  {item['table']} basis={item['basis']} rows="+
                     ', '.join(f"{r['ordinal']}[{r['proposed_role']}, active={r['active']}, "
                               f"class={r['class']}, test_hint={r['structured_test_hint']}]" for r in item['rows']))
    lines.append('\nZBIRNA NUMBER REFERENCES WITHOUT A CORE MATCH')
    for x in result['zbirna_number_not_found']:
        lines.append(f"  {x['source_table']} row={x['source_ordinal']} active={x['active']} {x['reason']}")
    lines.append('\nCORE TYPED ID REFERENCES:')
    fks=result['core_typed_fk_evidence']
    lines.append(f"  checked={fks['validated_fk_cells']} non_core_or_missing={fks['findings_total']}")
    for x in fks['by_type']:
        lines.append(f"  {x['finding']}: {x['source_table']}.{x['field']} -> {x['target_table']} x{x['count']}")
    lines.append('\nTERMINAL REFERENCES INTO OMITTED DOMAINS (must be cleared or normalized in new model)')
    for k,n in result['terminal_references_to_omitted_domains'].items():
        lines.append(f'  {k}: {n} rows')
    lines.append('\nSECONDARY FINANCE (not blockers for core conversion)')
    for name in sorted(FINANCE_SECONDARY):
        if name in counts:
            lines.append(f"  {name}: "+' '.join(f'{k}={v}' for k,v in sorted(counts[name].items())))
    lines.append('\nSUPPORT REVIEW (storno history, prices)')
    for name in sorted(SUPPORT_REVIEW):
        if name in counts:
            lines.append(f"  {name}: "+' '.join(f'{k}={v}' for k,v in sorted(counts[name].items())))
    lines.append('\nNON-CORE RECORDS ARE ARCHIVED IN THE ORIGINAL XLSM, NOT DELETED.')
    lines.append('REVIEW GENERACIJA/STORNO AND AMB OWNERSHIP BEFORE IMPORT.')
    return '\n'.join(lines)+'\n'


def main(argv=None)->int:
    p=argparse.ArgumentParser(description='MIG-01.4 read-only core migration scope')
    p.add_argument('--source',required=True)
    p.add_argument('--out',required=True)
    p.add_argument('--scope',default=str(Path(__file__).with_name('agrix_scope_operator_proposal.json')))
    a=p.parse_args(argv)
    source=Path(a.source).expanduser().resolve()
    out=Path(a.out).expanduser().resolve()
    if source.suffix.lower()!='.xlsm' or not source.is_file() or out==source.parent or out==source:
        p.error('Izvor mora biti .xlsm; izlazni direktorijum mora biti zaseban')
    try:
        sha=sha256_file(source)
        scope=json.loads(Path(a.scope).read_text(encoding='utf-8'))
        tabs=parse_xlsm(source)
        result=evaluate(tabs,scope,sha)
        out.mkdir(parents=True,exist_ok=True)
        (out/'agrix_core_scope_diagnostics.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        message=human_report(result)
        (out/'agrix_core_scope_diagnostics.txt').write_text(message,encoding='utf-8')
        print(message)
        print('Saved: agrix_core_scope_diagnostics.json and .txt')
        return 0
    except (AuditError,ValueError,KeyError,OSError,ET.ParseError,zipfile.BadZipFile) as e:
        print('MIG-01.4 STOP: '+type(e).__name__+' (no source workbook changes).',file=sys.stderr)
        return 2

if __name__=='__main__':
    sys.exit(main())
