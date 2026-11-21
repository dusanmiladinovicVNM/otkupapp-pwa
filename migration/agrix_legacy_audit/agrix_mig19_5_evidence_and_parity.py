#!/usr/bin/env python3
"""AGRIx MIG-19.5 -- exact legacy/AMB10 reconciliation evidence. READ ONLY.

Only purpose: find which recorded legacy legs account for the seven change points
and categorize otkup document-slice mismatches WITHOUT changing data or making
synthetic balancing entries. Stage and target are fixed/pinned snapshots.

Dependencies distributed in the ZIP: agrix_mig19_3_parity.py and
agrix_mig19_malina_readonly_audit.py.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import sys

from agrix_mig19_3_parity import (
    XlsmReadOnly, sha256, key, number, norm_type, date, void, dec,
    stage_rows, ledger_views, reconcile, write_csv, STAGE_SHA,
)
from agrix_mig19_4_difference_isolation import day_isolation, gate_sum, SIX, PRIOR_TARGET_SHA

CUTOFF = '2026-08-31'
PACK = '0,4KG'
EXPECTED_STATION_GAPS = {
    ('ST-00001', PACK): Decimal('-18'),
    ('ST-00002', PACK): Decimal('-83'),
    ('ST-00004', PACK): Decimal('500'),
    ('ST-00005', PACK): Decimal('-550'),
}


def endpoint(r, name):
    return (key(r.get(name+'NalogTip')), key(r.get(name+'NalogID')))


def old_signed(r):
    direction = key(r.get('Smer'))
    if direction not in {'ULAZ', 'IZLAZ'}:
        raise ValueError(f"Invalid legacy direction {direction} for AmbID {r.get('AmbID')}")
    q = number(r.get('Kolicina'), 'legacy quantity')
    return q if direction == 'ULAZ' else -q


def account_station_components(r, station):
    # A canonical single two-party move may touch the same named station only once.
    a, b = endpoint(r,'Od'),endpoint(r,'Na')
    q = number(r.get('Kolicina'), 'canonical quantity', positive=True)
    result=Decimal(0)
    if a==('STANICA',station):result -= q
    if b==('STANICA',station):result += q
    return result


def changes_by_station(old, projected, gaps):
    changes, first, _ = day_isolation(old, projected, gaps)
    return changes,first


def observe_source_and_target(old_rows, new_rows, changes):
    """Local private forensic rows; never emitted verbatim on stdout."""
    relevant={(c['NalogID'],c['TipAmbalaze'],c['Datum']) for c in changes if c['NalogTip']=='STANICA'}
    olddetail=[]; newdetail=[]
    bylegacy=defaultdict(list); bycanon=defaultdict(list)
    for n,r in enumerate(old_rows,1):
        day=date(r.get('Datum'),f'old#{n}');pack=norm_type(r.get('TipAmbalaze'))
        if void(r.get('Stornirano')) or day>CUTOFF:continue
        if key(r.get('EntitetTip'))!='STANICA':continue
        station=key(r.get('EntitetID'))
        k=(station,pack,day)
        if k not in relevant:continue
        signed=old_signed(r)
        item=dict(LegacyOrdinal=n,LegacyAmbID=key(r.get('AmbID')),
            StationID=station,Datum=day,TipAmbalaze=pack,
            Smer=key(r.get('Smer')),Kolicina=dec(abs(signed)),PotpisaniEfekat=dec(signed),
            DokumentTIP=key(r.get('DokumentTIP')),DokumentID=key(r.get('DokumentID')),
            VozacID=key(r.get('VozacID')))
        olddetail.append(item)
        bylegacy[k].append(item)
    for n,r in enumerate(new_rows,1):
        day=date(r.get('Datum'),f'new#{n}');pack=norm_type(r.get('TipAmbalaze'))
        if day>CUTOFF or key(r.get('VrstaKretanja'))=='ULAZ_TUDJE_AMBALAZE':continue
        a,b=endpoint(r,'Od'),endpoint(r,'Na')
        for st in set([a[1] if a[0]=='STANICA' else '',b[1] if b[0]=='STANICA' else '']):
            if not st:continue
            k=(st,pack,day)
            if k not in relevant:continue
            signed=account_station_components(r,st)
            item=dict(CanonicalOrdinal=n,CanonicalAmbID=key(r.get('AmbID')),
                StationID=st,Datum=day,TipAmbalaze=pack,PotpisaniEfekat=dec(signed),
                OdNalogTip=a[0],OdNalogID=a[1],NaNalogTip=b[0],NaNalogID=b[1],
                Kolicina=dec(number(r.get('Kolicina'),f'new#{n}')),
                VrstaKretanja=key(r.get('VrstaKretanja')),
                DokumentTIP=key(r.get('DokumentTIP')),DokumentID=key(r.get('DokumentID')),
                StornoOd=key(r.get('StornoOd')))
            newdetail.append(item)
            bycanon[k].append(item)
    breakdown=[]
    for c in changes:
        k=(c['NalogID'],c['TipAmbalaze'],c['Datum'])
        ol=bylegacy[k]; nl=bycanon[k]
        oldsum=sum((Decimal(x['PotpisaniEfekat']) for x in ol),Decimal(0))
        newsum=sum((Decimal(x['PotpisaniEfekat']) for x in nl),Decimal(0))
        if oldsum!=Decimal(c['StariDnevniEfekat']) or newsum!=Decimal(c['NoviDnevniEfekat']):
            raise ValueError('Breakdown does not independently reproduce daily changes: '+str(k))
        breakdown.append({**c,'LegacyRowsOnDate':len(ol),'CanonicalRowsOnDate':len(nl),
            'LegacyPositiveRows':sum(Decimal(x['PotpisaniEfekat'])>0 for x in ol),
            'LegacyNegativeRows':sum(Decimal(x['PotpisaniEfekat'])<0 for x in ol),
            'CanonicalPositiveRows':sum(Decimal(x['PotpisaniEfekat'])>0 for x in nl),
            'CanonicalNegativeRows':sum(Decimal(x['PotpisaniEfekat'])<0 for x in nl)})
    return olddetail,newdetail,breakdown


def known_cases(old_rows,new_rows,changes):
    """Test six user-confirmed stage source rows + historical ordinal 700.

    Don't assert old row 700 contains 500 until its actual fields prove it.
    A candidate same-date/quantity match in canonical is NOT proof of lineage.
    """
    old_byid=defaultdict(list)
    for n,r in enumerate(old_rows,1):old_byid[key(r.get('AmbID'))].append((n,r))
    six_results=[]
    all_subjects=list(SIX.items())
    for aid,(expected_station,expected_qty,expected_day) in all_subjects:
        hits=old_byid.get(aid,[])
        actual=[]
        for ordinal,r in hits:
            pack=norm_type(r.get('TipAmbalaze'))
            day=date(r.get('Datum'),f'old#{ordinal}')
            q=number(r.get('Kolicina'),f'old#{ordinal}')
            actual.append(dict(legacy_ordinal=ordinal,pack=pack,day=day,qty=q,
              station=key(r.get('EntitetID')),direction=key(r.get('Smer')),
              entity_type=key(r.get('EntitetTip')),stornirano=void(r.get('Stornirano'))))
        verified=len(actual)==1 and (actual[0]['pack'],actual[0]['day'],actual[0]['qty'],actual[0]['station'],actual[0]['direction'],actual[0]['entity_type'],actual[0]['stornirano']) == (
            PACK,expected_day,expected_qty,expected_station,'ULAZ','STANICA',False)
        matching=[]
        if verified:
            for n,t in enumerate(new_rows,1):
                if date(t.get('Datum'),f'new#{n}')!=expected_day or norm_type(t.get('TipAmbalaze'))!=PACK:continue
                if number(t.get('Kolicina'),f'new#{n}')!=expected_qty:continue
                if endpoint(t,'Na')==('STANICA',expected_station) and not key(t.get('StornoOd')):
                    matching.append((n,key(t.get('AmbID')),endpoint(t,'Od')))
        six_results.append(dict(Case=aid,StationID=expected_station,Datum=expected_day,
            Qty=dec(expected_qty),LegacyRows=len(hits),SourceSignatureVerified=verified,
            CanonicalSameDestinationCandidates=len(matching),
            CandidateCanonicalIDs=' | '.join(x[1] for x in matching),
            CandidateCounterparties=' | '.join(f'{x[2][0]}/{x[2][1]}' for x in matching),
            ExactLineageProven=False))
    row700=old_rows[699] if len(old_rows)>=700 else None
    qcase=None
    if row700 is not None:
        qcase=dict(ordinal=700,ambid=key(row700.get('AmbID')),date=date(row700.get('Datum'),'legacy ordinal 700'),
            pack=norm_type(row700.get('TipAmbalaze')),station=key(row700.get('EntitetID')),
            entity_type=key(row700.get('EntitetTip')),direction=key(row700.get('Smer')),
            quantity=dec(number(row700.get('Kolicina'),'legacy ordinal 700')),
            stornirano=void(row700.get('Stornirano')),
            doctype=key(row700.get('DokumentTIP')),docid=key(row700.get('DokumentID')))
    # Expected sole station ST-00004 daily difference: delta (new-old) = +500.
    gap500=[c for c in changes if c['NalogID']=='ST-00004' and c['TipAmbalaze']==PACK and c['Datum']=='2026-06-22']
    signature=(qcase is not None and qcase['date']=='2026-06-22' and qcase['pack']==PACK
               and qcase['entity_type']=='STANICA' and qcase['station']=='ST-00004'
               and qcase['direction']=='IZLAZ' and qcase['quantity']=='500'
               and not qcase['stornirano'] and key(qcase['doctype'])=='OTPREMNICA')
    if qcase is not None:qcase['MatchesST04Minus500Hypothesis']=signature
    return six_results,qcase, bool(gap500 and gap500[0]['PromenaRazlike']=='500')


def docs_by_key(olddoc,newdoc):
    """Group by doc/day/type with all participant deltas. No fabricated matching."""
    groups=defaultdict(list)
    for k in set(olddoc)|set(newdoc):
        a=olddoc[k];b=newdoc[k]
        if a==b:continue
        t,account,pack,day,doctype,docid=k
        if doctype!='OTKUP':continue
        groups[(day,pack,doctype,docid)].append(dict(AccountType=t,AccountID=account,
            OldSigned=dec(a),NewSigned=dec(b),Delta=dec(b-a)))
    full=[];classes=Counter();balanced_groups=0;groups_two_sides=0
    for (day,pack,doctype,docid),entries in sorted(groups.items()):
        tset={e['AccountType'] for e in entries}
        delta=sum((Decimal(e['Delta']) for e in entries),Decimal(0))
        label=('PAIR_COOP_STATION' if {'KOOPERANT','STANICA'}.issubset(tset) else 'OTHER')
        if len(entries)==2 and label=='PAIR_COOP_STATION':groups_two_sides+=1
        if delta==0:balanced_groups+=1
        classes[(label,'ZERO_NET' if delta==0 else 'NONZERO_NET',len(entries))]+=1
        full.append(dict(Datum=day,TipAmbalaze=pack,DokumentTIP=doctype,DokumentID=docid,
            AffectedAccounts=len(entries),AccountTypes=' + '.join(sorted(tset)),
            SumDelta=dec(delta),Classification=label,
            RowDeltas=' ; '.join(f"{e['AccountType']}/{e['AccountID']}:{e['Delta']}" for e in entries)))
    shareable=[dict(Classification=k[0],NetClass=k[1],AffectedAccountsPerDoc=k[2],Documents=v)
               for k,v in sorted(classes.items())]
    return full,shareable,dict(documents_with_otkup_delta=len(groups),
         documents_with_exactly_one_kooperant_one_station_delta=groups_two_sides,
         document_groups_whose_participant_deltas_sum_zero=balanced_groups,
         warning='Document net zero is NOT per-document account parity or event-lineage proof')


def synthetic_naive_projection(gaps):
    """Mathematical caution: adding three ST1->station moves breaks ST1 parity.

    Not a proposal to write; illustrates conservation obstacle only.
    """
    d={(r['NalogID'],r['TipAmbalaze']): Decimal(r['Razlika']) for r in gaps}
    planned=[('ST-00001',Decimal('5')),('ST-00001',Decimal('10')),
             ('ST-00001',Decimal('3')),('ST-00002',Decimal('83')),
             ('ST-00005',Decimal('300')),('ST-00005',Decimal('250'))]
    for s,q in planned:d[(s,PACK)] += q
    # In addition, if station 2 and 5 entries are minted as internal
    # ST1 -> destination transfers, ST1 is debited 633.
    d[('ST-00001',PACK)]-=Decimal(633)
    # Conditional 500 at ST4 is not applied here (source evidence must be checked)
    return dict(BlindSixMovesAndThreeInternalTransfers_ST1Delta=dec(d[('ST-00001',PACK)]),
                ST2Delta=dec(d[('ST-00002',PACK)]),ST5Delta=dec(d[('ST-00005',PACK)]),
                ST4DeltaUnchanged=dec(d[('ST-00004',PACK)]),
                why='Naively appending six rows plus ST1->ST2/5 transfers shifts 633 to ST1; this is NOT an approved migration plan')


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--stage',required=True,type=Path)
    parser.add_argument('--workbook',required=True,type=Path)
    parser.add_argument('--out-dir',required=True,type=Path)
    parser.add_argument('--expected-target-sha',default=PRIOR_TARGET_SHA)
    args=parser.parse_args(argv)
    src=args.stage/'UNCONVERTED__tblAmbalaza.jsonl'
    for f in (src,args.workbook):
        if not f.is_file():parser.error('Missing required input: '+str(f))
    out=args.out_dir.resolve()
    protected={src.resolve(),args.workbook.resolve()}
    if out in protected or out in {p.parent for p in protected}:parser.error('Output must be separate from input locations')
    before={str(p):sha256(p) for p in protected}
    if before[str(src.resolve())]!=STAGE_SHA:raise RuntimeError('Frozen staging SHA mismatch')
    if before[str(args.workbook.resolve())]!=args.expected_target_sha:raise RuntimeError('Pinned MIG19.3/19.4 target workbook SHA mismatch')
    legacy=stage_rows(src)
    reader=XlsmReadOnly(args.workbook)
    try:
        current=reader.read_table('tblAmbalaza',{'AmbID','Datum','TipAmbalaze','Kolicina',
            'OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja',
            'StornoOd','DokumentTIP','DokumentID'})
    finally:reader.close()
    old,physical,projected,cover,olddoc,newdoc,diag,after=ledger_views(legacy,current)
    account,dates,document=reconcile(old,physical,projected,cover,olddoc,newdoc)
    gaps,_=gate_sum(account)
    changes,first=changes_by_station(old,projected,gaps)
    olddetail,newdetail,breakdown=observe_source_and_target(legacy,current,changes)
    known,q700,qgap500=known_cases(legacy,current,changes)
    docpriv,docclass,docstats=docs_by_key(olddoc,newdoc)
    gap_map={(r['NalogID'],r['TipAmbalaze']):Decimal(r['Razlika']) for r in gaps}
    expected=(gap_map==EXPECTED_STATION_GAPS)
    seven=(len(changes)==7 and len(gaps)==4 and len(known)==6)
    proof6=all(k['SourceSignatureVerified'] for k in known)
    # For the six leg effects, signatures prove their presence in OLD, but not absence
    # from NEW; MIG19.4 candidate matches can be coincidental.
    end_sha={str(p):sha256(p) for p in protected}
    if before!=end_sha:raise RuntimeError('Inputs mutated during read-only audit')
    out.mkdir(parents=True,exist_ok=True)
    write_csv(out/'MIG195_SEVEN_CHANGEPOINTS_PRIVATE.csv',breakdown,
        ['NalogTip','NalogID','TipAmbalaze','Datum','StariDnevniEfekat',
        'NoviDnevniEfekat','PromenaRazlike','RazlikaOdPocetka',
        'LegacyRowsOnDate','CanonicalRowsOnDate','LegacyPositiveRows',
        'LegacyNegativeRows','CanonicalPositiveRows','CanonicalNegativeRows'])
    write_csv(out/'MIG195_LEGACY_CHANGEPOINT_ROWS_PRIVATE.csv',olddetail,
        ['LegacyOrdinal','LegacyAmbID','StationID','Datum','TipAmbalaze',
        'Smer','Kolicina','PotpisaniEfekat','DokumentTIP','DokumentID','VozacID'])
    write_csv(out/'MIG195_CANONICAL_CHANGEPOINT_ROWS_PRIVATE.csv',newdetail,
        ['CanonicalOrdinal','CanonicalAmbID','StationID','Datum','TipAmbalaze',
         'PotpisaniEfekat','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID',
         'Kolicina','VrstaKretanja','DokumentTIP','DokumentID','StornoOd'])
    write_csv(out/'MIG195_OMULAZ_SIX_CASES_PRIVATE.csv',known,
        ['Case','StationID','Datum','Qty','LegacyRows','SourceSignatureVerified',
         'CanonicalSameDestinationCandidates','CandidateCanonicalIDs',
         'CandidateCounterparties','ExactLineageProven'])
    if q700 is not None:
        (out/'MIG195_LEGACY_ORDINAL_700_PRIVATE.json').write_text(
            json.dumps(q700,ensure_ascii=False,indent=2),encoding='utf-8')
    write_csv(out/'MIG195_OTKUP_DOCUMENT_GROUPS_PRIVATE.csv',docpriv,
        ['Datum','TipAmbalaze','DokumentTIP','DokumentID','AffectedAccounts',
         'AccountTypes','SumDelta','Classification','RowDeltas'])
    write_csv(out/'MIG195_OTKUP_DOC_CLASSES_SHAREABLE.csv',docclass,
        ['Classification','NetClass','AffectedAccountsPerDoc','Documents'])
    total_delta=sum((Decimal(c['PromenaRazlike']) for c in changes),Decimal(0))
    summary={
        'format':'AGRIx-MIG19.5-FAITHFUL-EVIDENCE-v1',
        'status':'READ_ONLY_EVIDENCE_READY_NOT_MIGRATION_PASS',
        'target_sha256':before[str(args.workbook.resolve())],
        'stage_sha256_verified':True,'inputs_unchanged':True,
        'station_final_mismatches':len(gaps),
        'seven_change_point_pattern_verified':seven,
        'known_four_station_vector_verified':expected,
        'station_daily_change_points':len(changes),
        'station_change_summary':[dict(StationID=r['NalogID'],Datum=r['Datum'],
              TipAmbalaze=r['TipAmbalaze'],OldMinusNewDailyDelta=dec(-Decimal(r['PromenaRazlike'])))
              for r in sorted(changes,key=lambda r:(r['Datum'],r['NalogID']))],
        'six_omulaz_source_signatures_all_verified':proof6,
        'six_omulaz_canonical_destination_candidates_total':sum(int(r['CanonicalSameDestinationCandidates']) for r in known),
        'six_omulaz_proven_new_event_lineage':False,
        'legacy_source_ordinal_700':(None if q700 is None else {k:v for k,v in q700.items() if k!='docid'}),
        'st00004_500_daily_gap_is_present':qgap500,
        'st00004_500_is_source_ordinal_700':bool(q700 and q700['MatchesST04Minus500Hypothesis']),
        'document_groups':docstats,
        'document_classes':docclass,
        'sum_canonical_minus_legacy_station_change_points':dec(total_delta),
        'blind_append_conservation_warning':synthetic_naive_projection(gaps) if expected else None,
        'source_row_to_canonical_event_lineage_proven':False,
        'no_write_or_fix_applied':True,
        'conclusion':'Inspect local 7-date SOURCE/TARGET rows for *replacement* or exact provenance before any writer; no synthetic balancing entries'
    }
    (out/'MIG195_SUMMARY.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(summary,ensure_ascii=False,indent=2))
    return 0


if __name__=='__main__':
    try:sys.exit(main())
    except Exception as exc:
        print(f'MIG195_ERROR: {type(exc).__name__}: {exc}',file=sys.stderr)
        sys.exit(3)
