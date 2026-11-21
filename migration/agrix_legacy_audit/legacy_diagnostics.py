#!/usr/bin/env python3
"""AgriX MIG-01.1: second-pass read-only domain diagnostics of legacy .xlsm.

Uses the same stdlib OOXML reader as legacy_audit.py. No VBA, Excel, saving to
source, raw record values, names, addresses, bank details or real IDs in output.
Reports anonymous table ROW ORDINALS (within ListObject body, 1-based) and
per-run HMAC aliases to make exceptions traceable without exposing identities.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import json
from pathlib import Path
import sys
import zipfile
from xml.etree import ElementTree as ET

from legacy_audit import (AuditError, Inspection, Table, find_table, is_active,
                          num, parse_xlsm, sha256_file, trim)

SOURCE_SHA = '0687a7933c2a67cec71d6fc09180f46803dc6cc76af7db45cbe6cd74097265ab'


def row_view(index: int, row: dict) -> dict:
    return {'table_row_ordinal': index, 'active': is_active(row)}


def group_key(table: str, row: dict) -> tuple:
    metadata = {
        'tblOtkup': ('StanicaID', 'BrojDokumenta', ('KooperantID', 'Datum', 'KulturaID')),
        'tblOtpremnica': ('StanicaID', 'BrojOtpremnice', ('Datum', 'VozacID')),
        'tblZbirna': ('VozacID', 'BrojZbirne', ('Datum', 'KupacID')),
    }
    owner, docnum, attrs = metadata[table]
    gen = trim(row.get('GeneracijaID'))
    return ('GEN', gen) if gen else ('HEUR', trim(row.get(owner)), trim(row.get(docnum)),
                                      *(trim(row.get(a)) for a in attrs))


def diagnostics(tables: dict[str, Table]) -> dict:
    i = Inspection(tables)
    output = {
        'format': 'MIG01.1-privacy-safe',
        'privacy': 'NO_RAW_SOURCE_VALUES_OR_IDENTIFIERS; row ordinals + per-run anonymous HMAC aliases only',
        'basis': 'All findings are diagnostic, not proof of corrupt business data',
    }
    documents = {}
    for table in ('tblOtkup', 'tblOtpremnica', 'tblZbirna'):
        t = find_table(tables, table)
        if not t:
            continue
        groups = defaultdict(list)
        for idx, row in enumerate(t.rows, 1):
            groups[group_key(table, row)].append((idx, row))
        exceptions = []
        for key, entries in groups.items():
            classes = Counter(trim(r.get('Klasa')).upper() for _, r in entries
                              if trim(r.get('Klasa')))
            if max(classes.values(), default=0) > 1:
                exceptions.append({
                    'candidate_alias': i.alias(repr(key), table + '.candidate'),
                    'grouping_basis': key[0],
                    'classes': dict(classes),
                    'rows': [row_view(idx, r) for idx, r in entries],
                })
        documents[table] = {
            'source_rows': len(t.rows),
            'candidate_groups': len(groups),
            'suspected_duplicate_class_groups': exceptions,
            'active_rows': sum(is_active(r) for r in t.rows),
            'stornirane_rows': sum(not is_active(r) for r in t.rows),
        }
    output['documents'] = documents

    zb = find_table(tables, 'tblZbirna')
    if zb:
        zb_by_number = defaultdict(list)
        for idx, row in enumerate(zb.rows, 1):
            number = trim(row.get('BrojZbirne'))
            if number:
                zb_by_number[number].append((idx, row))
        links = {}
        for tname in ('tblOtkup', 'tblOtpremnica', 'tblPrijemnica'):
            tab = find_table(tables, tname)
            if not tab or 'BrojZbirne' not in tab.columns:
                continue
            missing, same_candidate, distinct_candidates_by_number, resolved_single = [], [], [], 0
            for idx, row in enumerate(tab.rows, 1):
                number = trim(row.get('BrojZbirne'))
                if not number:
                    continue
                matches = zb_by_number.get(number, [])
                if not matches:
                    missing.append({**row_view(idx, row),
                                    'zbirna_business_number_alias': i.alias(number, 'zbirna.number')})
                else:
                    candidates = {group_key('tblZbirna', zr) for _, zr in matches}
                    if len(candidates) == 1:
                        if len(matches) > 1:
                            same_candidate.append({**row_view(idx, row),
                                'zbirna_business_number_alias': i.alias(number, 'zbirna.number'),
                                'source_zbirna_rows': [zidx for zidx, _ in matches],
                                'source_zbirna_active': [is_active(zr) for _, zr in matches]})
                        else:
                            resolved_single += 1
                    else:
                        distinct_candidates_by_number.append({**row_view(idx, row),
                            'zbirna_business_number_alias': i.alias(number, 'zbirna.number'),
                            'candidate_group_count': len(candidates),
                            'source_zbirna_rows': [zidx for zidx, _ in matches],
                            'source_zbirna_active': [is_active(zr) for _, zr in matches]})
            links[tname] = {
                'not_found_by_number': len(missing),
                'missing_rows': missing,
                'multiple_rows_single_candidate_group': len(same_candidate),
                'multirow_same_group_examples': same_candidate[:30],
                'multiple_distinct_candidate_groups_by_number': len(distinct_candidates_by_number),
                'multiple_group_examples': distinct_candidates_by_number[:30],
                'exactly_one_row_by_number': resolved_single,
                'note': 'Even a single candidate by number is NOT validated ownership; cross-document identity validation is still mandatory',
            }
        output['zbirna_links'] = links

    finance = {}
    n = find_table(tables, 'tblNovac')
    if n:
        types = defaultdict(list)
        for idx, row in enumerate(n.rows, 1):
            label = trim(row.get('Tip'))
            # Only publish explicitly known, fixed code-defined types; never
            # free-form values, which could contain personal data.
            known = {'VirmanAvansKoop', 'VirmanFirmaKoop', 'KesFirmaOtkupac',
                     'KesOtkupacKoop', 'VirmanFirmaOtkupac', 'KupciUplata',
                     'KupciAvans', 'BankaUplata', 'BankaIsplata'}
            key = label if label in known else ('(prazno)' if not label else '(DRUGI_TIP)')
            types[key].append((idx,row))
        finance['tblNovac_by_type'] = {kind: {
            'rows': len(entries),
            'active_rows': sum(is_active(r) for _,r in entries),
            'stornirane_rows': sum(not is_active(r) for _,r in entries),
            'linked_otkup_rows': sum(bool(trim(r.get('OtkupID'))) for _,r in entries),
            'linked_faktura_rows': sum(bool(trim(r.get('FakturaID'))) for _,r in entries),
            'has_amount_rows': sum((num(r.get('Uplata')) or 0)!=0 or
                                   (num(r.get('Isplata')) or 0)!=0 for _,r in entries),
            'examples': [{**row_view(idx,r),
                          'has_otkup_id':bool(trim(r.get('OtkupID'))),
                          'has_faktura_id':bool(trim(r.get('FakturaID'))),
                          'has_kooperant_id':bool(trim(r.get('KooperantID')))}
                         for idx,r in entries[:12]],
        } for kind,entries in sorted(types.items())}
    o = find_table(tables, 'tblOtkup')
    if o:
        paid = [(idx,r) for idx,r in enumerate(o.rows,1) if
                trim(r.get('Isplaceno')).lower() in ('da','true','1')]
        money = [(idx,r) for idx,r in enumerate(o.rows,1) if (num(r.get('Novac')) or 0)!=0]
        finance['legacy_otkup_flags'] = {
            'paid_flag_rows':len(paid),'paid_active_rows':sum(is_active(r) for _,r in paid),
            'paid_example_ordinals':[idx for idx,_ in paid[:30]],
            'nonzero_novac_rows':len(money),
            'nonzero_novac_example_ordinals':[idx for idx,_ in money[:30]],
        }
    b = find_table(tables, 'tblBankaImport')
    if b:
        finance['tblBankaImport'] = {
            'rows': len(b.rows),
            'active_rows': sum(is_active(row) for row in b.rows),
            'columns': b.columns,
            'nonempty_by_column': {col:sum(bool(trim(row.get(col))) for row in b.rows)
                                   for col in b.columns},
            'note':'Bank staging content and bookkeeping identity are not inferred from row count.',
        }
    output['finance'] = finance

    t = find_table(tables, 'tblMagacin')
    if t:
        other = [(idx,r) for idx,r in enumerate(t.rows,1) if
                 trim(r.get('ArtikalID')) != 'ART-POC-DUG']
        opening = [(idx,r) for idx,r in enumerate(t.rows,1) if
                   trim(r.get('ArtikalID')) == 'ART-POC-DUG']
        output['agro'] = {
            'opening_rows': len(opening),
            'opening_active': sum(is_active(r) for _,r in opening),
            'other_rows':len(other),
            'other_records':[{
                **row_view(idx,r),
                'movement_type': trim(r.get('Tip')) if trim(r.get('Tip')) in ('Ulaz','Izlaz') else '(DRUGO)',
                'article_alias': i.alias(r.get('ArtikalID', ''), 'magacin.artikal'),
                'has_cooperator': bool(trim(r.get('KooperantID'))),
                'quantity_valid_positive': ((num(r.get('Kolicina')) or 0)>0),
                'value_valid': num(r.get('Vrednost')) is not None,
            } for idx,r in other],
            'opening_unique_cooperators':len(set(trim(r.get('KooperantID')) for _,r in opening if trim(r.get('KooperantID')))),
        }

    t = find_table(tables, 'tblAmbalaza')
    if t:
        doc_types = defaultdict(list)
        missing = []
        for idx, r in enumerate(t.rows,1):
            raw = trim(r.get('DokumentTip',r.get('DokumentTIP','')))
            category = ('(prazno)' if not raw else raw if raw in {
                'Otkup','Otpremnica','Zbirna','Prijemnica','Revers','AmbalazaDokument',
                'Povrat','Izdavanje','Ulaz','Izlaz'} else '(DRUGO:' + i.alias(raw,'amb.dokument_tip') + ')')
            doc_types[category].append((idx,r))
            if not trim(r.get('DokumentID')):
                missing.append({**row_view(idx,r),
                    'document_type_category': category,
                    'account_type': trim(r.get('EntitetTip')) if trim(r.get('EntitetTip')) in
                    ('Kooperant','Stanica','Kupac','Vozac','Firma','SpoljniSvet') else '(DRUGO)',
                    'has_driver':bool(trim(r.get('VozacID'))),
                    'quantity_positive':(num(r.get('Kolicina')) or 0)>0})
        output['ambalaza'] = {
            'types':{k:{'rows':len(v),'active':sum(is_active(r) for _,r in v),
                        'stornirane':sum(not is_active(r) for _,r in v)}
                     for k,v in sorted(doc_types.items())},
            'missing_document_rows':missing,
            'nonempty_unknown_type_kinds':sum(k.startswith('(DRUGO:') for k in doc_types),
            'note':'Unknown document categories are hashed; never interpret them as invalid merely because not in MIG-01 whitelist.',
        }
    return output


def format_report(d:dict) -> str:
    out = ['AGRIX LEGACY MIG-01.1 - SECOND-PASS READ-ONLY DIAGNOSTICS', '='*63,
           'No raw IDs, names, business numbers, money amounts or bank details.\n']
    for t, info in d.get('documents', {}).items():
        out.append(f"{t}: {info['source_rows']} rows; {info['candidate_groups']} candidate groups; "
                   f"duplicate-class groups={len(info['suspected_duplicate_class_groups'])}")
        for e in info['suspected_duplicate_class_groups']:
            out.append('  suspected class collision: ' + e['candidate_alias'] +
                       ', rows=' + str([r['table_row_ordinal'] for r in e['rows']]) +
                       ', classes=' + str(e['classes']))
    out.append('\nZBIRNA REFERENCES BY BUSINESS NUMBER (NOT VERIFIED IDS)')
    for t, data in d.get('zbirna_links', {}).items():
        out.append(f"{t}: missing={data['not_found_by_number']}, "
                   f"multirow-same-candidate={data['multiple_rows_single_candidate_group']}, "
                   f"multiple-distinct-candidates={data['multiple_distinct_candidate_groups_by_number']}")
        for e in data['missing_rows']:
            out.append('  missing: table row=' + str(e['table_row_ordinal']) +
                       ', active=' + str(e['active']) +
                       ', number=' + e['zbirna_business_number_alias'])
        for e in data['multiple_group_examples']:
            out.append('  ambiguous scope: table row=' + str(e['table_row_ordinal']) +
                       ', number=' + e['zbirna_business_number_alias'] +
                       ', candidate_groups=' + str(e['candidate_group_count']))
    out.append('\nFINANCE')
    for typ, data in d.get('finance', {}).get('tblNovac_by_type', {}).items():
        out.append(f"{typ}: count={data['rows']}, active={data['active_rows']}, "
                   f"storno={data['stornirane_rows']}, linked_otkup={data['linked_otkup_rows']}")
    out.append('Otkup flags: '+json.dumps(d.get('finance', {}).get('legacy_otkup_flags',{})))
    out.append('Bank import: '+json.dumps({k:v for k,v in d.get('finance', {}).get('tblBankaImport', {}).items()
                                           if k not in ('columns','nonempty_by_column')},ensure_ascii=False))
    out.append('\nAGRO')
    ag=d.get('agro',{})
    out.append('Opening='+str(ag.get('opening_rows'))+', other='+str(ag.get('other_rows')))
    for e in ag.get('other_records',[]):
        out.append(f"  other: row={e['table_row_ordinal']} active={e['active']} "
                   f"type={e['movement_type']} art={e['article_alias']}")
    out.append('\nAMBALAZA')
    for kind,info in d.get('ambalaza',{}).get('types',{}).items():
        out.append(f"{kind}: {info['rows']} rows (active={info['active']})")
    out.append('Missing document: '+json.dumps(d.get('ambalaza',{}).get('missing_document_rows',[])))
    out.append('\nCAVEAT: anonymized row ordinals refer to populated Excel ListObject rows; '
               'business decisions require source workbook review and explicit mapping rules.')
    return '\n'.join(out)+'\n'


def main(argv=None):
    ap=argparse.ArgumentParser(description='Second-pass privacy-safe read-only legacy diagnostics')
    ap.add_argument('--source',required=True)
    ap.add_argument('--out',required=True)
    ap.add_argument('--expect-sha256',default=SOURCE_SHA)
    args=ap.parse_args(argv)
    source=Path(args.source).expanduser().resolve()
    dest=Path(args.out).expanduser().resolve()
    if source.suffix.lower()!='.xlsm' or dest==source or dest==source.parent:
        ap.error('Izvor mora biti .xlsm, a output zaseban direktorijum')
    try:
        sh=sha256_file(source)
        if args.expect_sha256 and args.expect_sha256.lower()!=sh.lower():
            print('STOP: SHA256 stare sveske se razlikuje od MIG-01; nikakvi izvestaji nisu upisani.',file=sys.stderr)
            return 2
        result=diagnostics(parse_xlsm(source))
        result['source_sha256']=sh
        dest.mkdir(parents=True,exist_ok=True)
        (dest/'agrix_legacy_diagnostics.json').write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        text_report=format_report(result)
        (dest/'agrix_legacy_diagnostics.txt').write_text(text_report,encoding='utf-8')
    except (AuditError, OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        print('AUDIT FAILED: '+type(exc).__name__,file=sys.stderr)
        return 2
    print(text_report)
    print('Reports saved: agrix_legacy_diagnostics.json / agrix_legacy_diagnostics.txt')
    return 0

if __name__=='__main__':
    sys.exit(main())
