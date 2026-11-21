#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""AGRIx MIG-17.3 v1.1: read-only trace of one historic payment.

Source tblOtkup.Isplaceno may be a status flag/text, not a monetary value.
Such values are captured as unresolved evidence, NEVER silently treated as 0.
Do not change source, XLSM, Git or finance records. Send only summary JSON.
"""
import argparse
import json
import re
import unicodedata
from collections import Counter
from decimal import Decimal
from pathlib import Path

from agrix_mig171_reconcile import read_book, sha, norm, dec, day

AMOUNT = Decimal('100000')
TARGET = 'Milosavljevic Dragan'


def key(s):
    s = unicodedata.normalize('NFKD', str(s or '')).casefold()
    return ' '.join(''.join(c for c in s if not unicodedata.combining(c)).split())


def pick_stage(stage, filename):
    matches = list(stage.rglob(filename))
    if len(matches) != 1:
        raise ValueError('Expected exactly one stage file ' + filename + ', found ' + str(len(matches)))
    return matches[0]


def jsonl(p):
    with p.open('r', encoding='utf-8-sig') as f:
        for line in f:
            if line.strip():
                yield json.loads(line)


def legacy_payment_value(value):
    """Return (Decimal amount | None, classification).

    - Blank is NOT a positive payment.
    - Standard JSON numerics remain exact.
    - EU thousands pattern is interpreted only when it is explicit.
    - Text (DA/NE/status/etc.) is UNKNOWN, not a monetary zero.
    """
    if value is None or str(value).strip() == '':
        return Decimal(0), 'blank'
    if isinstance(value, bool):
        return None, 'non_numeric'
    raw = str(value).strip().replace('\xa0', ' ')
    stripped = raw.replace(' ', '')
    # e.g. 100.000,00 or 1.234.567; grouped decimal convention unambiguous.
    if re.fullmatch(r'[+-]?\d{1,3}(?:\.\d{3})+,\d{1,4}', stripped):
        return Decimal(stripped.replace('.', '').replace(',', '.')), 'numeric_eu'
    # e.g. 100 000,00 or 100 000; explicit space grouping.
    if re.fullmatch(r'[+-]?\d{1,3}(?: \d{3})+(?:,\d{1,4})?', raw):
        return Decimal(stripped.replace(',', '.')), 'numeric_eu'
    try:
        return dec(value, 'Otkup.Isplaceno'), 'numeric'
    except ValueError:
        return None, 'non_numeric'


def positive_payment(rows, person):
    return [r for r in rows
            if norm(r.get('KooperantID')) == norm(person)
            and dec(r.get('Isplata'), 'Novac.Isplata') > 0
            and norm(r.get('Stornirano')) != 'da']


def candidate_old_otkup(rows, person):
    """Return (positive numeric candidates, unknown nonnumeric source evidence, counts).

    Must not silently zero-out unknown source values. No business effect here.
    """
    paid = []
    unknown = []
    counts = Counter()
    for r in rows:
        if norm(r.get('KooperantID')) != norm(person) or norm(r.get('Stornirano')) == 'da':
            continue
        amount, kind = legacy_payment_value(r.get('Isplaceno'))
        counts[kind] += 1
        if amount is None:
            unknown.append(r)
        elif amount > 0:
            paid.append((r, amount))
    return paid, unknown, dict(counts)


def safe_dates(rows, field):
    count = Counter()
    for r in rows:
        raw = r.get(field)
        if raw is None or str(raw).strip() == '':
            count['blank'] += 1
        else:
            try:
                _ = day(raw)
                count['valid'] += 1
            except (ValueError, TypeError):
                count['invalid'] += 1
    return dict(count)


def run(proj, xlsm, outdir, station):
    proj = proj.resolve()
    xlsm = xlsm.resolve()
    outdir = outdir.resolve()
    if outdir.exists():
        raise ValueError('Output dir already exists; choose a fresh --out-dir')
    if not xlsm.is_file():
        raise ValueError('Input .xlsm missing')

    stage = proj / 'MigrationStage_v2'
    before = sha(xlsm)
    book, _ = read_book(xlsm, ['tblNovac', 'tblKooperanti', 'tblOtkup'])
    if len(book['tblNovac']) != 39:
        raise ValueError('Expected canonical 39 novac rows')

    target = key(TARGET)
    people = [r for r in book['tblKooperanti']
              if key(str(r.get('Prezime', '')) + ' ' + str(r.get('Ime', ''))) == target
              or key(str(r.get('Ime', '')) + ' ' + str(r.get('Prezime', ''))) == target]
    if len(people) != 1:
        raise ValueError('Target name must match exactly ONE tblKooperanti row, found ' + str(len(people)))
    person = people[0]
    pid = str(person.get('KooperantID') or '').strip()
    if not pid:
        raise ValueError('Target row has no KooperantID')

    old_otkup = list(jsonl(pick_stage(stage, 'SOURCE__tblOtkup.jsonl')))
    old_novac = list(jsonl(pick_stage(stage, 'SECONDARY__tblNovac.jsonl')))
    if len(old_novac) != 39:
        raise ValueError('Source secondary financial row count unexpected')

    old_paid, old_unknown, parse_counts = candidate_old_otkup(old_otkup, pid)
    new_paid = positive_payment(book['tblNovac'], pid)
    old_sec_paid = positive_payment(old_novac, pid)

    legacy_amount = sum((v for r, v in old_paid), Decimal(0))
    new_amount = sum((dec(r.get('Isplata')) for r in new_paid), Decimal(0))
    finance_same_amount = [r for r in book['tblNovac'] if dec(r.get('Isplata')) == AMOUNT]

    def brief(rows, fields):
        return [{k: r.get(k) for k in fields} for r in rows]

    private = {
        'subject_kooperant_id': pid,
        'subject_name': str(person.get('Ime', '')) + ' ' + str(person.get('Prezime', '')),
        'legacy_otkup_numeric_paid_candidates': [
            {**{k: r.get(k) for k in ('OtkupID', 'Datum', 'DatumIsplate', 'Isplaceno', 'Stornirano', 'KooperantID', 'StanicaID')},
             'interpreted_amount': str(v)} for r, v in old_paid],
        'legacy_otkup_NONNUMERIC_isplaceno_evidence': brief(old_unknown,
            ['OtkupID', 'Datum', 'DatumIsplate', 'Isplaceno', 'Novac', 'Stornirano', 'KooperantID', 'StanicaID']),
        'legacy_otkup_isplaceno_parse_counts': parse_counts,
        'canonical_novac_subject_rows': brief([
            r for r in book['tblNovac'] if norm(r.get('KooperantID')) == norm(pid)],
            ['NovacID', 'Datum', 'Isplata', 'Uplata', 'Tip', 'KooperantID', 'PartnerID', 'OMID', 'OtkupID', 'Stornirano']),
        'old_secondary_novac_subject_rows': brief([
            r for r in old_novac if norm(r.get('KooperantID')) == norm(pid)],
            ['NovacID', 'Datum', 'Isplata', 'Uplata', 'KooperantID', 'PartnerID', 'OMID', 'OtkupID', 'Stornirano']),
        'all_existing_100000_payment_rows': brief(finance_same_amount,
            ['NovacID', 'Datum', 'KooperantID', 'PartnerID', 'OMID', 'OtkupID', 'Tip', 'Isplata']),
        'legacy_POSITIVE_NUMERIC_candidate_total_NOT_CERTIFIED': str(legacy_amount),
        'new_novac_positive_paid_sum': str(new_amount),
        'interpretation_warning': 'Legacy Otkup.Isplaceno may be a paid/not-paid status, not money. Non-numeric rows must be manually interpreted, not zeroed.'
    }
    if sha(xlsm) != before:
        raise ValueError('Input XLSM unexpectedly changed during reading')

    summary = {
        'format': 'AGRIx-MIG173-TARGETED-PAYMENT-TRACE-v1.1',
        'status': 'READ_ONLY_TRACE_WITH_UNRESOLVED_SOURCE_FIELD' if old_unknown else 'READ_ONLY_TRACE_COMPLETE',
        'subject': 'ONE_100000_RSD_OLD_REPORT_PAYMENT',
        'input_sha256': before,
        'workbook_untouched': True,
        'git_untouched': True,
        'source_legacy_otkup_rows': len(old_otkup),
        'source_secondary_novac_rows': len(old_novac),
        'target_resolved_unique': True,
        'target_maticna_om_matches_station': norm(person.get('StanicaID')) == norm(station),
        'legacy_otkup_numeric_positive_paid_candidates': len(old_paid),
        'legacy_otkup_non_numeric_isplaceno_count': len(old_unknown),
        'legacy_otkup_isplaceno_value_class_counts': parse_counts,
        'legacy_otkup_numeric_candidates_total_100000': legacy_amount == AMOUNT,
        'legacy_otkup_payment_total_proven': False,
        'legacy_otkup_numeric_candidate_has_100000_single_row': any(v == AMOUNT for _, v in old_paid),
        'legacy_otkup_numeric_candidate_payment_dates': safe_dates([r for r, _ in old_paid], 'DatumIsplate'),
        'secondary_novac_subject_positive_rows': len(old_sec_paid),
        'canonical_novac_subject_positive_rows': len(new_paid),
        'canonical_novac_subject_has_100000': any(dec(r.get('Isplata')) == AMOUNT for r in new_paid),
        'canonical_novac_any_100000_rows': len(finance_same_amount),
        'source_legacy_otkup_isplaceno_is_only_a_candidate': 'No evidence that source Isplaceno is a monetary amount; raw nonnumeric values retained privately.',
        'private_file_created': True,
        'contains_private_customer_values': False,
        'next_step': 'Interpret target old payment fields with ledger; do not append money automatically.'
    }

    outdir.mkdir(parents=True)
    (outdir / 'MIG173_missing_payment_PRIVATE.json').write_text(
        json.dumps(private, indent=2, ensure_ascii=False, default=str) + '\n', encoding='utf-8')
    (outdir / 'MIG173_summary.json').write_text(
        json.dumps(summary, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    print('MIG-17.3 v1.1 READ-ONLY TRACE:', summary['status'])
    print('Historical numeric candidates:', len(old_paid),
          'uninterpreted Isplaceno fields:', len(old_unknown))
    print('SHARE ONLY:', outdir / 'MIG173_summary.json')
    print('PRIVATE file remains local')
    return summary


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--project-dir', required=True, type=Path)
    p.add_argument('--input-xlsm', required=True, type=Path)
    p.add_argument('--out-dir', required=True, type=Path)
    p.add_argument('--station-id', default='ST-00001')
    a = p.parse_args()
    run(a.project_dir, a.input_xlsm, a.out_dir, a.station_id)


if __name__ == '__main__':
    main()
