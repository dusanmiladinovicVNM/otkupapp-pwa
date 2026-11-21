#!/usr/bin/env python3
"""MIG-01.2: read-only evidence for separating legacy customer and test rows.

Does not guess or remove test data. Optionally analyzes human-specified table
row cutoffs and checks exact ID references across the proposed boundary.
No raw workbook values, identifiers, amounts, names or dates leave process memory.
Only OOXML read, never Excel/VBA or workbook save.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

from legacy_audit import AuditError, Table, parse_xlsm, sha256_file, trim

SOURCE_SHA = '0687a7933c2a67cec71d6fc09180f46803dc6cc76af7db45cbe6cd74097265ab'
TOKEN_RE = re.compile(r'(?:^|[^A-Z0-9])(?:TST|SVT|BIT|GLD|TEST)(?:[^A-Z0-9]|$)', re.I)
SOFT_RE = re.compile(r'\b(?:test|testni|testna|testno|proba|probni|probna|demo|fixture)\b', re.I)
KNOWN_SEED_VALUES = {'ST-90001', 'KOOP-90001'}  # public repo's test protocol, advisory only
DESC_COLS = {
    'naziv', 'ime', 'prezime', 'mesto', 'adresa', 'partner', 'opis',
    'napomena', 'svrhaplacanja', 'komentar', 'nazivfirme', 'nazivpartnera',
}
PK_BY_TABLE = {
    'tblAmbalaza': 'AmbID', 'tblArtikli': 'ArtikalID',
    'tblBankaImport': 'BankaImportID', 'tblCenovnik': 'CenovnikID',
    'tblFakture': 'FakturaID', 'tblFakturaStavke': 'FakturaStavkaID',
    'tblKooperanti': 'KooperantID', 'tblKulture': 'KulturaID',
    'tblKupci': 'KupacID', 'tblMagacin': 'MagacinID',
    'tblNovac': 'NovacID', 'tblOtkup': 'OtkupID',
    'tblOtpremnica': 'OtpremnicaID', 'tblPaleta': 'PaletaID',
    'tblPaletaStavka': 'PaletaStavkaID', 'tblParcele': 'ParcelaID',
    'tblPrerada': 'PreradaID', 'tblPreradaStavka': 'PreradaStavkaID',
    'tblPrijemnica': 'PrijemnicaID', 'tblPrevoznici': 'PrevoznikID',
    'tblStanice': 'StanicaID', 'tblStornoVeze': 'StornoVezaID',
    'tblStornoZurnal': 'StornoZurnalID', 'tblUtovar': 'UtovarID',
    'tblUtovarStavke': 'UtovarStavkaID', 'tblVozaci': 'VozacID',
    'tblZbirna': 'ZbirnaID',
}


def cell_signals(column: str, value: str) -> set[str]:
    """Signals are evidence, never proof; no raw value returned."""
    val = trim(value)
    if not val:
        return set()
    upper = val.upper()
    col = column.lower()
    if upper in KNOWN_SEED_VALUES:
        return {'repo_test_seed'}
    # 'test' in a real person's name is never enough to exclude.
    if col in DESC_COLS:
        return {'descriptive_test_word'} if SOFT_RE.search(val) else set()
    if col.endswith('id') or col.startswith('broj') or col in (
        'vrstavoca', 'sortavoca', 'tipambalaze', 'kultura',
        'clientrecordid', 'generacijaid',
    ):
        return {'structured_test_token'} if TOKEN_RE.search(val) else set()
    return set()


def scan_row(row: dict[str, str]) -> tuple[set[str], list[str]]:
    all_signals: set[str] = set()
    columns: list[str] = []
    for col, val in row.items():
        found = cell_signals(col, val)
        if found:
            all_signals.update(found)
            columns.append(col)
    return all_signals, columns


def classify_tables(tables: dict[str, Table], tail_window: int = 25) -> tuple[dict, dict]:
    report = {}
    signals_by_table: dict[str, dict[int, set[str]]] = {}
    for _key, table in sorted(tables.items()):
        n = len(table.rows)
        items: dict[int, set[str]] = {}
        hits = []
        for ordinal, row in enumerate(table.rows, 1):
            signals, columns = scan_row(row)
            if signals:
                items[ordinal] = signals
                hits.append({'ordinal': ordinal, 'signal_classes': sorted(signals),
                             'columns': columns})
        signals_by_table[table.name] = items
        head = max(1, n - tail_window + 1)
        report[table.name] = {
            'row_count': n,
            'strong_signal_rows': sum(bool(v & {'repo_test_seed','structured_test_token'}) for v in items.values()),
            'weak_signal_rows': sum(bool(v & {'descriptive_test_word'}) for v in items.values()),
            'strong_first_ordinal': next((h['ordinal'] for h in hits if
                          set(h['signal_classes']) & {'repo_test_seed','structured_test_token'}), None),
            'strong_last_ordinal': next((h['ordinal'] for h in reversed(hits) if
                          set(h['signal_classes']) & {'repo_test_seed','structured_test_token'}), None),
            'tail_window': min(n, tail_window),
            'tail_window_signal_rows': [h for h in hits if h['ordinal'] >= head],
            'earlier_signal_ordinals': [h['ordinal'] for h in hits if h['ordinal'] < head][:40],
            'earlier_signals_truncated': max(0, sum(h['ordinal'] < head for h in hits) - 40),
            'scope_decision': 'NOT_CLASSIFIED',
        }
    return report, signals_by_table


def scope_template(tables: dict[str, Table], sha: str) -> dict:
    return {
        'source_sha256': sha,
        'operation': 'PROPOSED_TEST_SCOPE_ONLY_DO_NOT_DELETE',
        'exclude_tail_from': {t.name: None for _name, t in sorted(tables.items())},
        'instructions': ('Set 1-based first TEST row per table only after examining original '
                         'workbook. 1=all rows test; null=no excluded rows. '
                         'Cross-links in next report must be resolved before migration.'),
    }


def validate_scope(scope: dict, tables: dict[str, Table], sha: str) -> dict[str, int]:
    if scope.get('source_sha256') != sha:
        raise AuditError('STOP: scope source_sha256 se ne poklapa sa sveskom')
    allowed = {table.name.lower(): table for table in tables.values()}
    provided = scope.get('exclude_tail_from')
    if not isinstance(provided, dict):
        raise AuditError('scope.exclude_tail_from mora biti objekat')
    boundaries = {}
    for name, cutoff in provided.items():
        if name.lower() not in allowed:
            raise AuditError('Nepostojeca tabela u scope-u (proveri lokalni scope)')
        table = allowed[name.lower()]
        if cutoff is None:
            continue
        if isinstance(cutoff, bool) or not isinstance(cutoff, int):
            raise AuditError('Granica mora biti celobrojni 1-based ordinal ili null')
        if not 1 <= cutoff <= len(table.rows):
            raise AuditError('Granica je van opsega postojecih popunjenih redova')
        boundaries[table.name] = cutoff
    return boundaries


def boundary_report(tables: dict[str, Table], boundaries: dict[str, int],
                    signals_by_table: dict[str, dict[int,set[str]]]) -> dict:
    """Check exact ID crossings. A crossing is a review item, not evidence of corruption."""
    pk_value_index = defaultdict(list)
    for table in tables.values():
        pk = PK_BY_TABLE.get(table.name)
        if pk and pk in table.columns:
            for idx, row in enumerate(table.rows, 1):
                key = trim(row.get(pk))
                if key:
                    pk_value_index[key].append((table.name, idx))
    details = []
    cross_counts: Counter = Counter()
    seen = set()
    # Lower bound on cross-boundary relationships: exact matching foreign IDs.
    for table in tables.values():
        source_pk = PK_BY_TABLE.get(table.name)
        cut = boundaries.get(table.name, len(table.rows)+1)
        for i, row in enumerate(table.rows, 1):
            source_excluded = i >= cut
            for col, val in row.items():
                if not col.lower().endswith('id') or col == source_pk:
                    continue
                key = trim(val)
                if not key:
                    continue
                for dst, j in pk_value_index.get(key, []):
                    if (table.name,i)==(dst,j):
                        continue
                    dst_cut = boundaries.get(dst, len(tables[dst.lower()].rows)+1)
                    target_excluded = j >= dst_cut
                    if source_excluded == target_excluded:
                        continue
                    id_key = (table.name,i,col,dst,j)
                    if id_key in seen:
                        continue
                    seen.add(id_key)
                    kind = ('RETAINED_REFERENCES_EXCLUDED' if not source_excluded else
                            'EXCLUDED_REFERENCES_RETAINED')
                    cross_counts[(kind,table.name,col,dst)] += 1
                    if len(details) < 150:
                        details.append({'kind':kind,'from_table':table.name,'from_row_ordinal':i,
                                        'foreign_key_column':col,'to_table':dst,'to_row_ordinal':j})
    counts = {}
    for table in tables.values():
        n = len(table.rows)
        cut = boundaries.get(table.name, n+1)
        if cut <= n:
            kept, excluded = cut-1, n-cut+1
            strong_after = sum(bool(v & {'repo_test_seed','structured_test_token'})
                               for ordinal,v in signals_by_table[table.name].items() if ordinal >= cut)
            strong_before = sum(bool(v & {'repo_test_seed','structured_test_token'})
                                for ordinal,v in signals_by_table[table.name].items() if ordinal < cut)
            counts[table.name]={'retained':kept,'proposed_excluded':excluded,
                                'structured_signals_retained':strong_before,
                                'structured_signals_proposed_excluded':strong_after}
    return {
        'table_counts':counts,
        'cross_boundary_id_references':sum(cross_counts.values()),
        'cross_boundary_by_type':[{'kind':kind,'from_table':src,'field':col,
                                   'to_table':dst,'count':count}
                                  for (kind,src,col,dst),count in sorted(cross_counts.items())],
        'cross_boundary_examples':details,
        'examples_truncated':max(0,sum(cross_counts.values())-len(details)),
        'limits': ('Only exact foreign-ID matches covered. Business-number-based links, '
                   'polymorphic event ownership and synthetic IDs need domain review. '
                   'Test scope is not approved for migration.'),
    }


def human_report(data: dict) -> str:
    lines=['AGRIX MIG-01.2 - TEST/PODACI SCOPE REVIEW (READ ONLY)', '='*60,
           'No raw customer values, IDs, names, amounts or bank details.',
           'Signatures are HINTS, never grounds for automatic deletion.',
           f"Source SHA256: {data['source_sha256']}", '']
    for name, d in sorted(data['tables'].items()):
        lines.append(f"{name:29} rows={d['row_count']:5} strict_hints={d['strong_signal_rows']:4} "
                     f"soft_hints={d['weak_signal_rows']:4} first={d['strong_first_ordinal']} "
                     f"last={d['strong_last_ordinal']}")
        if d['tail_window_signal_rows']:
            samples=d['tail_window_signal_rows'][-15:]
            lines.append('  tail hinted ordinals: ' + ', '.join(
                str(x['ordinal'])+' ('+'/'.join(x['signal_classes'])+')' for x in samples))
    sc=data.get('proposed_scope')
    if sc:
        lines.append('\nPROPOSED SCOPE - DRY RUN ONLY')
        for name, v in sc['table_counts'].items():
            lines.append(f"  {name}: retained={v['retained']} proposed_excluded={v['proposed_excluded']} "
                         f"strict_retained={v['structured_signals_retained']}")
        lines.append(f"  Cross-boundary exact ID references: {sc['cross_boundary_id_references']}")
        for kind in sc['cross_boundary_by_type'][:50]:
            lines.append(f"    {kind['kind']} {kind['from_table']}.{kind['field']} -> "
                         f"{kind['to_table']}: {kind['count']}")
        lines.append('  BUSINESS NUMBER & AMB ownership links are NOT covered by this check.')
    lines.append('\nNO MIGRATION / NO SOURCE CHANGES / NO AUTOMATIC EXCLUSIONS')
    return '\n'.join(lines)+'\n'


def main(argv=None) -> int:
    ap=argparse.ArgumentParser(description='MIG-01.2: read-only test-data candidate isolation')
    ap.add_argument('--source',required=True)
    ap.add_argument('--out',required=True)
    ap.add_argument('--scope',help='optional manually edited scope JSON, for dry-run crosslinks only')
    ap.add_argument('--tail',type=int,default=25,help='number of tail rows to display marker hints for')
    ap.add_argument('--expect-sha256',default=SOURCE_SHA)
    a=ap.parse_args(argv)
    source=Path(a.source).expanduser().resolve()
    out=Path(a.out).expanduser().resolve()
    if source.suffix.lower()!='.xlsm' or not source.is_file() or out==source or out==source.parent:
        ap.error('Izvor mora biti postojeci .xlsm, --out mora biti zaseban folder')
    if not 1 <= a.tail <= 250:
        ap.error('--tail mora biti 1..250')
    if a.scope and Path(a.scope).resolve()==source:
        ap.error('scope nije isti fajl kao source')
    try:
        sha=sha256_file(source)
        if a.expect_sha256 and sha.lower()!=a.expect_sha256.lower():
            raise AuditError('Source SHA256 nije isti kao prethodni audit')
        tables=parse_xlsm(source)
        report, signs=classify_tables(tables,a.tail)
        result={'format':'MIG01.2','source_sha256':sha,
                'privacy':'AGGREGATES_ORDINALS_AND_COLUMN_NAMES_ONLY',
                'scope_status':'NO_PRODUCTION_TEST_CLASSIFICATION_YET',
                'tables':report}
        if a.scope:
            scope=json.loads(Path(a.scope).read_text(encoding='utf-8'))
            boundaries=validate_scope(scope,tables,sha)
            result['proposed_scope']=boundary_report(tables,boundaries,signs)
        out.mkdir(parents=True,exist_ok=True)
        (out/'agrix_test_scope_diagnostics.json').write_text(
            json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (out/'agrix_test_scope_diagnostics.txt').write_text(human_report(result),encoding='utf-8')
        template=out/'agrix_test_scope_template.json'
        if not template.exists():
            template.write_text(json.dumps(scope_template(tables,sha),indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        print(human_report(result))
        print('Saved: agrix_test_scope_diagnostics.json/.txt; template scope JSON')
        return 0
    except (AuditError, ValueError, OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as ex:
        print('MIG-01.2 FAILED: '+str(ex), file=sys.stderr)
        return 2


if __name__=='__main__':
    sys.exit(main())
