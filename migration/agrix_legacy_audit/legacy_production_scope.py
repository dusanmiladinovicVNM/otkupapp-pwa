#!/usr/bin/env python3
"""MIG-01.3: read-only, operator-specified candidate production/test scope.

No workbook writes, no macros or COM. No sensitive cell values leave memory.
Scope decisions are only PROPOSED until candidate crosslinks and review rows
are resolved. We do not delete, export or migrate ANY workbook rows.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET

from legacy_audit import AuditError, Table, parse_xlsm, sha256_file, trim, is_active
from legacy_diagnostics import group_key
from legacy_test_scope import classify_tables, PK_BY_TABLE

KEPT = 'KEEP_CANDIDATE'
TEST = 'TEST_CANDIDATE'
REVIEW = 'REVIEW_REQUIRED'
UNKNOWN = 'UNCLASSIFIED'

# Typed relationships only: matching arbitrary PK strings across tables can
# otherwise produce false links. 'DokumentID' is handled as polymorphic.
FK_TO_TARGET = {
    'KooperantID':'tblKooperanti', 'StanicaID':'tblStanice',
    'VozacID':'tblVozaci', 'KulturaID':'tblKulture',
    'ParcelaID':'tblParcele', 'OtkupID':'tblOtkup',
    'OtpremnicaID':'tblOtpremnica', 'ZbirnaID':'tblZbirna',
    'PrijemnicaID':'tblPrijemnica', 'FakturaID':'tblFakture',
    'KupacID':'tblKupci', 'ArtikalID':'tblArtikli',
    'PaletaID':'tblPaleta', 'PrevoznikID':'tblPrevoznici',
    'UtovarID':'tblUtovar', 'PreradaID':'tblPrerada',
    'MagacinID':'tblMagacin', 'NovacID':'tblNovac',
}
DOCUMENT_TABLES = ('tblOtkup', 'tblOtpremnica', 'tblZbirna', 'tblPrijemnica',
                   'tblFakture', 'tblPaleta', 'tblPrerada', 'tblUtovar')

def parse_scope(scope: dict, tables: dict[str, Table], sha: str) -> dict[str, list[str]]:
    if scope.get('source_sha256', '').lower() != sha.lower():
        raise AuditError('STOP: source SHA256 se ne poklapa sa predlogom scope-a')
    if scope.get('format') != 'MIG01.3-scope-proposal':
        raise AuditError('Nepodrzan format scope-a')
    raw_cutoffs = scope.get('last_real_row_inclusive', {})
    all_test = scope.get('all_test_tables', [])
    review_rows = scope.get('review_rows', {})
    review_all = scope.get('review_entire_tables', [])
    if not isinstance(raw_cutoffs, dict) or not isinstance(all_test, list) or not isinstance(review_rows, dict) or not isinstance(review_all, list):
        raise AuditError('Neispravan tip polja scope-a')
    by_name = {t.name.lower():t for t in tables.values()}
    for name in (*raw_cutoffs, *all_test, *review_rows, *review_all):
        if not isinstance(name, str) or name.lower() not in by_name:
            raise AuditError('Scope navodi tabelu koja ne postoji u staroj svesci')
    used = [name.lower() for name in raw_cutoffs] + [name.lower() for name in all_test] + [name.lower() for name in review_all]
    if len(set(used)) != len(used):
        raise AuditError('Tabela sme imati samo jedno osnovno pravilo')
    state = {}
    for table in tables.values():
        name = table.name
        n = len(table.rows)
        if name in raw_cutoffs:
            cutoff = raw_cutoffs[name]
            if isinstance(cutoff, bool) or not isinstance(cutoff, int) or not 0 <= cutoff <= n:
                raise AuditError('Neispravan poslednji poslovni red za '+name)
            current = [KEPT if i <= cutoff else TEST for i in range(1,n+1)]
        elif name in all_test:
            current = [TEST]*n
        elif name in review_all:
            current = [REVIEW]*n
        else:
            current = [UNKNOWN]*n
        state[name] = current
    for name, ordinals in review_rows.items():
        if not isinstance(ordinals, list) or not ordinals or len(set(map(str,ordinals)))!=len(ordinals):
            raise AuditError('Neispravni review redovi za '+name)
        for ordinal in ordinals:
            if isinstance(ordinal,bool) or not isinstance(ordinal,int) or not 1 <= ordinal <= len(state[name]):
                raise AuditError('Review ordinal van opsega: '+name)
            state[name][ordinal-1] = REVIEW
    return state


def table_summary(tables: dict[str,Table], state:dict[str,list[str]],
                  strong: dict[str,dict[int,set[str]]]) -> dict:
    report={}
    for table in sorted(tables.values(),key=lambda t:t.name):
        s=state[table.name]
        cnt=Counter(s)
        strong_kept=[i for i,sig in strong[table.name].items()
                     if s[i-1]==KEPT and sig & {'structured_test_token','repo_test_seed'}]
        strong_test=[i for i,sig in strong[table.name].items()
                     if s[i-1]==TEST and sig & {'structured_test_token','repo_test_seed'}]
        report[table.name]={
            'source_rows':len(s), 'keep_candidate':cnt[KEPT], 'test_candidate':cnt[TEST],
            'review_required':cnt[REVIEW], 'unclassified':cnt[UNKNOWN],
            'strong_test_marker_in_kept_count':len(strong_kept),
            'strong_test_marker_in_kept_examples':strong_kept[:20],
            'strong_test_marker_in_test_count':len(strong_test),
        }
    return report


def candidate_group_boundary(tables:dict[str,Table],state:dict[str,list[str]]) -> list[dict]:
    out=[]
    for name in ('tblOtkup','tblOtpremnica','tblZbirna'):
        tab=tables.get(name.lower())
        if not tab:continue
        groups=defaultdict(list)
        for i,row in enumerate(tab.rows,1):
            groups[group_key(name,row)].append(i)
        for key,indices in groups.items():
            states={state[name][i-1] for i in indices}
            if len(states)>1:
                out.append({'table':name,'grouping_basis':key[0],
                            'ordinals':indices,'status':[state[name][i-1] for i in indices],
                            'note':'Candidate group only; GEN more reliable than HEUR.'})
    return out


def exact_id_links(tables:dict[str,Table],state:dict[str,list[str]])->dict:
    indices={}
    for table in tables.values():
        pk=PK_BY_TABLE.get(table.name)
        if pk and pk in table.columns:
            d=defaultdict(list)
            for i,row in enumerate(table.rows,1):
                val=trim(row.get(pk))
                if val:d[val].append(i)
            indices[table.name]=d
    counts=Counter()
    examples=[]
    for table in tables.values():
        for i,row in enumerate(table.rows,1):
            a=state[table.name][i-1]
            if a not in (KEPT,REVIEW):continue
            for col,value in row.items():
                val=trim(value)
                if not val:continue
                dests=(DOCUMENT_TABLES if col=='DokumentID' else
                       (FK_TO_TARGET[col],) if col in FK_TO_TARGET else ())
                for target in dests:
                    for j in indices.get(target,{}).get(val,[]):
                        if table.name==target and i==j:continue
                        b=state[target][j-1]
                        if b==KEPT:continue
                        kind=f'{a}_TO_{b}'
                        counts[(kind,table.name,col,target)] += 1
                        if len(examples)<120:
                            examples.append({'kind':kind,'from_table':table.name,
                                             'from_row':i,'column':col,
                                             'to_table':target,'to_row':j})
    return {
      'count':sum(counts.values()),
      'by_type':[{'kind':kind,'from_table':src,'column':col,'to_table':dest,'count':n}
                 for (kind,src,col,dest),n in sorted(counts.items())],
      'examples':examples,
      'examples_truncated':max(0,sum(counts.values())-len(examples)),
      'limitation':'Typed explicit FK values + polymorphic DokumentID only. Does not prove business-number links or account ownership.'
    }


def zbirna_number_links(tables:dict[str,Table],state:dict[str,list[str]])->dict:
    zbirna=tables.get('tblzbirna')
    if not zbirna:return {}
    bynumber=defaultdict(list)
    for i,row in enumerate(zbirna.rows,1):
        key=trim(row.get('BrojZbirne'))
        if key:bynumber[key].append((i,row))
    result={}
    for name in ('tblOtkup','tblOtpremnica','tblPrijemnica'):
        tab=tables.get(name.lower())
        if not tab or 'BrojZbirne' not in tab.columns:continue
        counters=Counter()
        examples=[]
        for i,row in enumerate(tab.rows,1):
            if state[name][i-1]!=KEPT:continue
            number=trim(row.get('BrojZbirne'))
            if not number:continue
            candidates=bynumber.get(number,[])
            live=[(j,r) for j,r in candidates if state['tblZbirna'][j-1]==KEPT]
            if not candidates:kind='NOT_FOUND_ANYWHERE'
            elif not live:kind='ONLY_TEST_OR_REVIEW_CANDIDATES'
            else:
                candidate_groups={group_key('tblZbirna',r) for _,r in live}
                kind=('MULTIPLE_DISTINCT_KEPT_GROUPS' if len(candidate_groups)>1 else
                      'SINGLE_GROUP_WITH_MULTIPLE_ROWS' if len(live)>1 else 'ONE_KEPT_ROW')
            counters[kind]+=1
            if kind!='ONE_KEPT_ROW' and len(examples)<55:
                examples.append({'source_row':i,'status':kind,
                                 'candidate_zbirna_rows':[j for j,_ in candidates][:12],
                                 'candidate_zbirna_states':[state['tblZbirna'][j-1] for j,_ in candidates][:12],
                                 'active':is_active(row)})
        result[name]={'counts':dict(counters),'examples':examples,
                      'note':'BrojZbirne alone never authorizes ID linking.'}
    return result


def summarize(result:dict)->str:
    lines=['AGRIX MIG-01.3 - OPERATOR SCOPE / CROSS-BOUNDARY DIAGNOSTICS',
           '='*68, 'READ ONLY - PROPOSED SCOPE, NOT AUTHORIZATION TO MIGRATE',
           'No raw names, IDs, amounts, bank details, dates or document numbers.',
           '']
    for table,rec in result['tables'].items():
        if rec['source_rows'] == 0:continue
        lines.append(f"{table:25} total={rec['source_rows']:5} keep={rec['keep_candidate']:5} "
                     f"test={rec['test_candidate']:5} review={rec['review_required']:4} "
                     f"unknown={rec['unclassified']:5} hints_in_keep={rec['strong_test_marker_in_kept_count']}")
    lines.append('\nPOSSIBLE SPLIT DOCUMENT GROUPS ACROSS PRODUCTION/TEST SCOPE')
    for d in result['document_group_boundary'][:40]:
        lines.append(f"  {d['table']}: rows={d['ordinals']} states={d['status']} grouping={d['grouping_basis']}")
    lines.append(f"  Total groups crossing boundary: {len(result['document_group_boundary'])}")
    links=result['exact_id_links']
    lines.append(f"\nKEPT/REVIEW SOURCE -> NON-KEPT TARGET EXACT ID LINKS: {links['count']}")
    for d in links['by_type'][:35]:
        lines.append(f"  {d['kind']}: {d['from_table']}.{d['column']} -> {d['to_table']} x{d['count']}")
    lines.append('\nZBIRNA REFS IN KEPT DOCUMENT ROWS (BUSINESS NUMBER ONLY)')
    for name,rec in result['zbirna_number_links'].items():
        lines.append(f'  {name}: '+', '.join(f'{k}={v}' for k,v in sorted(rec['counts'].items())))
        for ex in rec['examples']:
            if ex['status'] in ('NOT_FOUND_ANYWHERE','ONLY_TEST_OR_REVIEW_CANDIDATES','MULTIPLE_DISTINCT_KEPT_GROUPS'):
                lines.append(f"    row {ex['source_row']} {ex['status']} candidate_rows={ex['candidate_zbirna_rows']}")
    lines.append('\nOPEN DECISIONS: REVIEW_REQUIRED rows and UNCLASSIFIED tables stay unresolved.')
    lines.append('NO EXCEL/VBA EXECUTED; NO WORKBOOK DATA WAS CHANGED OR EXPORTED.')
    return '\n'.join(lines)+'\n'


def evaluate(tables:dict[str,Table],scope:dict,sha:str)->dict:
    state=parse_scope(scope,tables,sha)
    _sigreport,signals=classify_tables(tables)
    return {'format':'MIG01.3-scope-assessment', 'source_sha256':sha,
            'scope_status':'PROPOSAL_NOT_APPROVED',
            'privacy':'ONLY TABLE NAMES, ENUM-LIKE STATUS, ROW ORDINALS, COUNTS',
            'tables':table_summary(tables,state,signals),
            'document_group_boundary':candidate_group_boundary(tables,state),
            'exact_id_links':exact_id_links(tables,state),
            'zbirna_number_links':zbirna_number_links(tables,state)}


def main(argv=None)->int:
    ap=argparse.ArgumentParser(description='MIG01.3 operator-defined scope review; read-only')
    ap.add_argument('--source',required=True,help='August customer .xlsm; never modified')
    ap.add_argument('--out',required=True,help='report folder (not source workbook folder)')
    ap.add_argument('--scope',default=str(Path(__file__).with_name('agrix_scope_operator_proposal.json')))
    args=ap.parse_args(argv)
    source=Path(args.source).expanduser().resolve()
    out=Path(args.out).expanduser().resolve()
    if source.suffix.lower()!='.xlsm' or not source.is_file() or out in (source, source.parent):
        ap.error('Izvor mora biti postojeci .xlsm, a --out zaseban folder')
    if Path(args.scope).resolve()==source:
        ap.error('Scope nikada ne sme biti stara sveska')
    try:
        scope=json.loads(Path(args.scope).read_text(encoding='utf-8'))
        sha=sha256_file(source)
        if sha.lower()!=str(scope.get('source_sha256','')).lower():
            raise AuditError('STOP: sveska ima drugi SHA256 od operatorovog predloga scope-a')
        tables=parse_xlsm(source)
        result=evaluate(tables,scope,sha)
        out.mkdir(parents=True,exist_ok=True)
        (out/'agrix_production_scope_diagnostics.json').write_text(
            json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        report=summarize(result)
        (out/'agrix_production_scope_diagnostics.txt').write_text(report,encoding='utf-8')
        print(report)
        print('Saved: agrix_production_scope_diagnostics.json/.txt')
        return 0
    except (AuditError,ValueError,OSError,KeyError,zipfile.BadZipFile,ET.ParseError):
        print('MIG-01.3 STOP: scope/workbook validation failed (source unchanged).',file=sys.stderr)
        return 2

if __name__=='__main__':
    sys.exit(main())
