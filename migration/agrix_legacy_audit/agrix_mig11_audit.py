#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIx MIG-11 — READ-ONLY audit istorijskih veza, standardna Python biblioteka.

Bez Excel COM-a, bez menjanja stage-a / xlsm / VBA / Git-a. Ne pokusava da
pretvori nejedinstveni poslovni broj u kanonski ID. Privatni zapisi ostaju lokalni.
"""
import argparse
import csv
import hashlib
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

CRITICAL = ('tblOtkup', 'tblOtpremnica', 'tblZbirna')
LINKS = ('tblOtpremnicaIzvori', 'tblZbirnaIzvori')
OPTIONAL = ('tblNovac', 'tblAmbalaza', 'tblAmbalazaDokument', 'tblStornoVeze')
TABLES = CRITICAL + LINKS + OPTIONAL
KEYS = {
    'tblOtkup': 'OtkupID',
    'tblOtpremnica': 'OtpremnicaID',
    'tblZbirna': 'ZbirnaID',
    'tblOtpremnicaIzvori': 'OtpremnicaIzvorID',
    'tblZbirnaIzvori': 'ZbirnaIzvorID',
}
FIELDS = {
    'tblOtkup': ['OtkupID', 'OtpremnicaID', 'BrojOtpremnice', 'StanicaID', 'Stornirano'],
    'tblOtpremnica': ['OtpremnicaID', 'ZbirnaRoditeljID', 'BrojOtpremnice', 'BrojZbirne', 'VozacID', 'Stornirano'],
    'tblZbirna': ['ZbirnaID', 'BrojZbirne', 'VozacID', 'Stornirano'],
    'tblOtpremnicaIzvori': ['OtpremnicaID', 'OtkupID'],
    'tblZbirnaIzvori': ['ZbirnaID', 'OtpremnicaID'],
}


def norm(x):
    if x is None:
        return ''
    return str(x).strip().casefold()


def is_cancelled(r):
    return norm(r.get('Stornirano')) in ('da', 'yes', 'true', '1')


def table_of_path(path):
    stem = path.stem.casefold()
    for name in sorted(TABLES, key=len, reverse=True):
        n = name.casefold()
        # exact or stage_export_<tblName>_mapped — but no accidental partial matches
        if re.search(r'(^|[^a-z0-9])' + re.escape(n) + r'($|[^a-z0-9])', stem):
            return name
        if stem == n:
            return name
    return None


def read_text(path):
    raw = path.read_bytes()
    if raw.startswith((b'\xff\xfe', b'\xfe\xff')):
        return raw.decode('utf-16')
    for enc in ('utf-8-sig', 'cp1250'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ValueError('Tekst nije podrzanog encodinga')


def read_rows(path):
    text = read_text(path)
    suffix = path.suffix.lower()
    if suffix in ('.csv', '.tsv'):
        delimiter = '\t' if suffix == '.tsv' else ','
        if suffix == '.csv':
            try:
                delimiter = csv.Sniffer().sniff(text[:8192], delimiters=',;\t').delimiter
            except csv.Error:
                delimiter = ','
        reader = csv.DictReader(text.splitlines(), delimiter=delimiter)
        if not reader.fieldnames:
            raise ValueError('Nema zaglavlja CSV')
        rows = []
        for row in reader:
            if not any(norm(v) for v in row.values() if v is not None):
                continue
            rows.append({str(k).strip().lstrip('\ufeff'): (v if v is not None else '') for k, v in row.items() if k is not None})
        return rows
    if suffix in ('.json', '.jsonl', '.ndjson'):
        if suffix in ('.jsonl', '.ndjson'):
            arr = [json.loads(x) for x in text.splitlines() if x.strip()]
        else:
            arr = json.loads(text)
        if isinstance(arr, dict):
            columns = arr.get('columns', [])
            arr = arr.get('rows', arr.get('records', arr.get('data', [])))
        else:
            columns = []
        if not isinstance(arr, list):
            raise ValueError('Nepoznat JSON raspored (ocekivani rows/lista)')
        if arr and isinstance(arr[0], list) and columns:
            arr = [dict(zip(columns, row)) for row in arr]
        if not all(isinstance(row, dict) for row in arr):
            raise ValueError('JSON redovi nisu objekti sa imenovanim kolonama')
        return [{str(k).strip().lstrip('\ufeff'): v for k, v in row.items()} for row in arr]
    raise ValueError('Nepodrzan tip fajla')


def inventory(stage):
    by_table = defaultdict(list)
    # Stage fajlovi su jedina osnova: preskoci zavrsne izvode i prethodne audit rezultate.
    for p in sorted(stage.rglob('*')):
        if not p.is_file() or p.suffix.lower() not in ('.csv', '.tsv', '.json', '.jsonl', '.ndjson'):
            continue
        if any(part.casefold() in ('reports', 'report', 'audit', 'diagnostics', 'logs') for part in p.parts):
            continue
        name = table_of_path(p)
        if name:
            by_table[name].append(p)
    return by_table


def load_stage(stage):
    inv = inventory(stage)
    loaded, errors, selected = {}, [], {}
    # MIG-09 cilj ima tacna kanonska imena: tblOtkup.jsonl itd.
    # SOURCE__/SECONDARY__/UNCONVERTED__ su evidencija porekla, NE ciljna tabela.
    for t in TABLES:
        files = inv.get(t, [])
        canonical = [p for p in files if p.stem.casefold() == t.casefold()]
        if len(canonical) > 1:
            errors.append(f'{t}: vise kanonskih fajlova ({len(canonical)}); prekinuto')
            loaded[t] = None
            continue
        if not canonical:
            loaded[t] = None
            if t in CRITICAL:
                errors.append(f'{t}: kanonski {t}.jsonl/.csv nije pronadjen')
            continue
        path = canonical[0]
        try:
            rows = read_rows(path)
            required = {'tblOtkup': ('OtkupID',),
                        'tblOtpremnica': ('OtpremnicaID',),
                        'tblZbirna': ('ZbirnaID',)}.get(t, ())
            if rows:
                present_cols = set().union(*(set(row) for row in rows))
                missing = [col for col in required if col not in present_cols]
                if missing:
                    raise ValueError('kanonske kolone nedostaju: ' + ', '.join(missing))
            loaded[t] = rows
            selected[t] = path.relative_to(stage).as_posix()
        except Exception as e:
            errors.append(f'{t}: ne mogu da procitam {path.name} ({type(e).__name__}: {e})')
            loaded[t] = None
    return loaded, inv, errors, selected


def index_rows(rows, key):
    result = defaultdict(list)
    for r in rows or []:
        result[norm(r.get(key))].append(r)
    return result


def evaluate(data):
    stats, detail = {}, []
    for t in TABLES:
        rows = data.get(t)
        stats[t] = {'present': rows is not None, 'rows': len(rows) if rows is not None else None}
        if rows is not None and t in KEYS:
            idx = index_rows(rows, KEYS[t])
            stats[t]['blank_ids'] = len(idx.get('', []))
            stats[t]['duplicate_distinct_ids'] = sum(1 for k, v in idx.items() if k and len(v) > 1)
    if any(data.get(t) is None for t in CRITICAL):
        return stats, detail, {'error': 'Nedostaju ili se ne mogu ucitati obavezne tri tabele'}

    otkups, otps, zbrs = (data[t] for t in CRITICAL)
    stage_cols = {t: set().union(*(set(r) for r in (data.get(t) or []))) for t in CRITICAL}
    for table, need in [('tblOtkup','OtpremnicaID'),('tblOtkup','BrojOtpremnice'),('tblOtpremnica','ZbirnaRoditeljID'),('tblOtpremnica','BrojZbirne')]:
        if need not in stage_cols[table]:
            # Nepostojeca kolona NIJE isto sto i prazan podatak u svakom redu.
            stats[table].setdefault('absent_optional_link_columns', []).append(need)
    oi, pi, zi = (index_rows(otkups, 'OtkupID'), index_rows(otps, 'OtpremnicaID'), index_rows(zbrs, 'ZbirnaID'))
    existing_o = defaultdict(set)
    existing_z = defaultdict(set)
    counters = Counter()
    conflicts = Counter()

    for row in data.get('tblOtpremnicaIzvori') or []:
        src, dst = norm(row.get('OtkupID')), norm(row.get('OtpremnicaID'))
        if not src or not dst or len(oi.get(src, [])) != 1 or len(pi.get(dst, [])) != 1:
            conflicts['existing_otpremnica_link_missing_or_invalid_fk'] += 1
            continue
        existing_o[src].add(dst)
    for row in data.get('tblZbirnaIzvori') or []:
        src, dst = norm(row.get('OtpremnicaID')), norm(row.get('ZbirnaID'))
        if not src or not dst or len(pi.get(src, [])) != 1 or len(zi.get(dst, [])) != 1:
            conflicts['existing_zbirna_link_missing_or_invalid_fk'] += 1
            continue
        existing_z[src].add(dst)
    conflicts['otkup_multi_existing_link_sources'] = sum(1 for v in existing_o.values() if len(v) > 1)
    conflicts['otpremnica_multi_existing_link_sources'] = sum(1 for v in existing_z.values() if len(v) > 1)

    otp_by_label = defaultdict(list)
    for r in otps:
        otp_by_label[(norm(r.get('StanicaID')), norm(r.get('BrojOtpremnice')))].append(r)
    zbr_by_label = defaultdict(list)
    for r in zbrs:
        zbr_by_label[(norm(r.get('VozacID')), norm(r.get('BrojZbirne')))].append(r)

    def add(typ, src, dst, classification, reasons):
        counters[f'{typ}:{classification}'] += 1
        detail.append({
            'link_type': typ, 'source_id': src, 'target_id': dst,
            'class': classification, 'reason': ';'.join(sorted(set(reasons)))
        })

    for r in otkups:
        src, dst = norm(r.get('OtkupID')), norm(r.get('OtpremnicaID'))
        if not src or len(oi.get(src, [])) != 1:
            add('OTKUP->OTPREMNICA', src, dst, 'INVALID_SOURCE', ['blank_or_duplicate_source_id'])
            continue
        if dst:
            reasons = []
            if len(pi.get(dst, [])) != 1:
                reasons.append('missing_or_duplicate_target_id')
            if is_cancelled(r):
                reasons.append('source_cancelled_historical')
            if dst in [k for k in existing_o.get(src, set())] and len(existing_o[src]) == 1:
                add('OTKUP->OTPREMNICA', src, dst, 'ALREADY_PRESENT', reasons)
                continue
            if existing_o.get(src) and dst not in existing_o[src]:
                reasons.append('conflicts_with_existing_membership')
            if len(pi.get(dst, [])) == 1:
                target = pi[dst][0]
                if is_cancelled(target):
                    reasons.append('target_cancelled_historical')
                if norm(r.get('StanicaID')) and norm(target.get('StanicaID')) and norm(r.get('StanicaID')) != norm(target.get('StanicaID')):
                    reasons.append('station_mismatch')
                if norm(r.get('BrojOtpremnice')) and norm(target.get('BrojOtpremnice')) and norm(r.get('BrojOtpremnice')) != norm(target.get('BrojOtpremnice')):
                    reasons.append('document_number_mismatch')
            add('OTKUP->OTPREMNICA', src, dst, 'EXPLICIT_ID_REVIEW' if reasons else 'EXPLICIT_ID_CANDIDATE', reasons)
        elif existing_o.get(src):
            add('OTKUP->OTPREMNICA', src, '', 'ALREADY_PRESENT', ['canonical_relation_only'])
        elif 'OtpremnicaID' not in stage_cols['tblOtkup'] and 'BrojOtpremnice' not in stage_cols['tblOtkup']:
            counters['OTKUP->OTPREMNICA:LINK_FIELDS_NOT_EXPORTED'] += 1
        elif norm(r.get('BrojOtpremnice')):
            label = (norm(r.get('StanicaID')), norm(r.get('BrojOtpremnice')))
            candidates = otp_by_label.get(label, [])
            cls = 'LABEL_ONLY_UNIQUE' if len(candidates) == 1 else ('LABEL_AMBIGUOUS' if candidates else 'LABEL_NOT_FOUND')
            add('OTKUP->OTPREMNICA', src, norm(candidates[0].get('OtpremnicaID')) if len(candidates) == 1 else '', cls, ['number_not_identity'])
        else:
            counters['OTKUP->OTPREMNICA:NO_REFERENCE'] += 1

    for r in otps:
        src, dst = norm(r.get('OtpremnicaID')), norm(r.get('ZbirnaRoditeljID'))
        if not src or len(pi.get(src, [])) != 1:
            add('OTPREMNICA->ZBIRNA', src, dst, 'INVALID_SOURCE', ['blank_or_duplicate_source_id'])
            continue
        if dst:
            reasons = []
            if len(zi.get(dst, [])) != 1:
                reasons.append('missing_or_duplicate_target_id')
            if is_cancelled(r):
                reasons.append('source_cancelled_historical')
            if dst in existing_z.get(src, set()) and len(existing_z[src]) == 1:
                add('OTPREMNICA->ZBIRNA', src, dst, 'ALREADY_PRESENT', reasons)
                continue
            if existing_z.get(src) and dst not in existing_z[src]:
                reasons.append('conflicts_with_existing_membership')
            if len(zi.get(dst, [])) == 1:
                target = zi[dst][0]
                if is_cancelled(target):
                    reasons.append('target_cancelled_historical')
                if norm(r.get('VozacID')) and norm(target.get('VozacID')) and norm(r.get('VozacID')) != norm(target.get('VozacID')):
                    reasons.append('driver_mismatch')
                if norm(r.get('BrojZbirne')) and norm(target.get('BrojZbirne')) and norm(r.get('BrojZbirne')) != norm(target.get('BrojZbirne')):
                    reasons.append('document_number_mismatch')
            add('OTPREMNICA->ZBIRNA', src, dst, 'EXPLICIT_ID_REVIEW' if reasons else 'EXPLICIT_ID_CANDIDATE', reasons)
        elif existing_z.get(src):
            add('OTPREMNICA->ZBIRNA', src, '', 'ALREADY_PRESENT', ['canonical_relation_only'])
        elif 'ZbirnaRoditeljID' not in stage_cols['tblOtpremnica'] and 'BrojZbirne' not in stage_cols['tblOtpremnica']:
            counters['OTPREMNICA->ZBIRNA:LINK_FIELDS_NOT_EXPORTED'] += 1
        elif norm(r.get('BrojZbirne')):
            key = (norm(r.get('VozacID')), norm(r.get('BrojZbirne')))
            matches = zbr_by_label.get(key, [])
            cls = 'LABEL_ONLY_UNIQUE' if len(matches) == 1 else ('LABEL_AMBIGUOUS' if matches else 'LABEL_NOT_FOUND')
            add('OTPREMNICA->ZBIRNA', src, norm(matches[0].get('ZbirnaID')) if len(matches) == 1 else '', cls, ['number_not_identity'])
        else:
            counters['OTPREMNICA->ZBIRNA:NO_REFERENCE'] += 1

    return stats, detail, {'candidate_counts': dict(sorted(counters.items())), 'conflicts': dict(sorted(conflicts.items()))}


def column_profile(paths, stage):
    """Bez privatnih vrednosti: samo imena kolona, broj redova i polja sa popunjenim vrednostima."""
    profiles = {}
    for t in CRITICAL + LINKS:
        profiles[t] = []
        for p in paths.get(t, []):
            try:
                rows = read_rows(p)
                counts = Counter(k for row in rows for k, v in row.items() if norm(v))
                fields = sorted(set(k for row in rows for k in row))
                profiles[t].append({'file': p.relative_to(stage).as_posix(), 'rows': len(rows),
                                    'columns': fields, 'nonempty_by_column': dict(sorted(counts.items()))})
            except Exception as ex:
                profiles[t].append({'file': p.relative_to(stage).as_posix(),
                                    'read_error': type(ex).__name__ + ': ' + str(ex)[:120]})
    return profiles


def sha256(path):
    h = hashlib.sha256()
    with path.open('rb') as fp:
        for chunk in iter(lambda: fp.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser(description='MIG-11 read-only stage link auditor')
    ap.add_argument('--stage', type=Path, required=True, help='lokalni MigrationStage_v2')
    ap.add_argument('--out', type=Path, required=True, help='lokalni output direktorijum (van stage)')
    args = ap.parse_args(argv)
    stage, out = args.stage.resolve(), args.out.resolve()
    if not stage.is_dir():
        ap.error(f'Ne postoji stage direktorijum: {stage}')
    if out == stage or stage in out.parents:
        ap.error('Output mora biti VAN source stage direktorijuma')
    if out.exists() and any(out.iterdir()):
        ap.error('Output direktorijum mora biti nov/prazan, da se nista ne prepise')

    data, inv, load_errors, selected = load_stage(stage)
    stats, detail, findings = evaluate(data)
    digest = {}
    for t, paths in inv.items():
        digest[t] = [{'name': p.relative_to(stage).as_posix(), 'sha256': sha256(p)} for p in paths]
    report = {
        'format': 'AGRIx-MIG-11-READ-ONLY-PREFLIGHT-v1.2',
        'execution_utc': datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'policy': 'No user data edited, no inferred links written; numbers are candidates only',
        'tables': stats,
        'files': digest,
        'selected_canonical_inputs': selected,
        'unconverted_or_source_only': {t: [p.relative_to(stage).as_posix() for p in paths] for t, paths in inv.items() if t not in selected},
        'load_errors': load_errors,
        'findings': findings,
        'private_candidate_rows': len(detail),
        'input_column_profiles': column_profile(inv, stage),
        'data_readiness': ('PREFLIGHT_INCOMPLETE' if findings.get('error') or load_errors
                           else 'AUDIT_WITH_MISSING_LINK_FIELDS' if (not ({'OtpremnicaID','BrojOtpremnice'} & set().union(*(set(r) for r in (data.get('tblOtkup') or [])))) or not ({'ZbirnaRoditeljID','BrojZbirne'} & set().union(*(set(r) for r in (data.get('tblOtpremnica') or [])))))
                           else 'AUDIT_COMPLETE'),
    }
    out.mkdir(parents=True, exist_ok=True)
    (out / 'mig11_summary.json').write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    with (out / 'mig11_candidates_PRIVATE_LOCAL.csv').open('w', encoding='utf-8-sig', newline='') as fh:
        writer = csv.DictWriter(fh, fieldnames=['link_type', 'source_id', 'target_id', 'class', 'reason'])
        writer.writeheader()
        writer.writerows(detail)
    print('MIG-11 READ-ONLY:', report['data_readiness'])
    print('Source is unchanged. No XLSM or VBA files opened/modified.')
    for t in CRITICAL + LINKS + OPTIONAL:
        row = stats[t]
        print(f'  {t}: {row["rows"] if row["present"] else "MISSING"}')
    if findings.get('candidate_counts'):
        for k, v in findings['candidate_counts'].items():
            print(f'  {k}: {v}')
    for err in load_errors:
        print('  INPUT_ERROR:', err)
    if findings.get('error'):
        print('  INPUT_ERROR:', findings['error'])
    print('Share only mig11_summary.json; NEVER share mig11_candidates_PRIVATE_LOCAL.csv')
    print('Out:', out)
    return 2 if findings.get('error') or load_errors else 0


if __name__ == '__main__':
    sys.exit(main())
