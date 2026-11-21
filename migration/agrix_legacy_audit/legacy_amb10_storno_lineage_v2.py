#!/usr/bin/env python3
"""MIG-04.6: read-only legacy storno/packaging lineage evidence (_v2 only).

A parent document marked storno while an AMB row remains active is a real state
mismatch. A same-amount reversal elsewhere is at most a candidate, NOT proof
of the original transaction being canceled. No target XLSM or AMB-10 entries.
Only local PRIVATE files show actual business identifiers and amounts.
"""
from __future__ import annotations
import argparse
from collections import Counter, defaultdict
from decimal import Decimal
import json
from pathlib import Path
import sys

from legacy_audit import AuditError, is_active, trim
from legacy_membership_v2 import get_rows
import legacy_amb10_forensics_v2 as f
from legacy_amb10_shadow_v2 import raw_contributions

VERSION='MIG-04.6'
F_FIELDS=['source_row_ordinal','AmbID','DokumentTip','DokumentID','Datum','TipAmbalaze',
          'Kolicina','Smer','EntitetTip','EntitetID','VozacID','Stornirano']
STORNO_TRUE={'da','true','1','yes'}
STORNO_FALSE={'','ne','false','0','no'}


def flag_state(r):
    value=f.key(r.get('Stornirano',''))
    if value not in STORNO_TRUE|STORNO_FALSE:
        return 'UNKNOWN'
    return 'ACTIVE' if is_active(r) else 'STORNO'


def contribution_vector(row):
    # Treat row as ACTIVE for movement-shape comparison, even when the source
    # history marked it Stornirano. Do not call this a booked balance.
    active=dict(row, Stornirano='')
    contributions=raw_contributions(active)
    d=defaultdict(Decimal)
    for (kind,eid,typ),amount in contributions:
        d[(f.key(kind),f.key(eid),f.key(typ))]+=amount
    return tuple(sorted((key,str(value)) for key,value in d.items() if value))


def negated_vector(vector):
    return tuple(sorted((k,str(-Decimal(v))) for k,v in vector))


def stage_records(stage):
    manifest,tables,idx=f.inputs(stage)  # source SHA, counts and needed columns guarded
    amb=tables['UNCONVERTED__tblAmbalaza']
    if len(amb)!=7815:raise AuditError('Unexpected legacy AMB row count')
    for j,r in enumerate(amb,1):
        if not trim(r.get('AmbID')):raise AuditError('Empty AmbID at ordinal '+str(j))
        if flag_state(r)=='UNKNOWN':
            raise AuditError('Unexpected Stornirano token in ambalaza at ordinal '+str(j))
    return manifest,tables,idx,amb


def build(stage,focus_ordinal=700,focus_id='AMB-00700'):
    manifest,tables,idx,source=stage_records(stage)
    if not 1<=focus_ordinal<=len(source):raise AuditError('Invalid focus ordinal')
    by_id={};doc_buckets=defaultdict(list);vectors=defaultdict(list)
    for i,r in enumerate(source,1):
        aid=trim(r['AmbID'])
        if aid in by_id:raise AuditError('Duplicate legacy AmbID near ordinal '+str(i))
        by_id[aid]=i
        doc_buckets[(f.key(r.get('DokumentTip')),trim(r.get('DokumentID')))].append(i)
        if flag_state(r)=='ACTIVE':
            v=contribution_vector(r)
            if v:vectors[v].append(i)
    if focus_id not in by_id:raise AuditError('Focus AmbID not found in pinned source')
    focus_id_ordinal=by_id[focus_id]
    focus_row=source[focus_ordinal-1]
    focus_same=(focus_id_ordinal==focus_ordinal)
    stats=Counter({k:0 for k in ('LINKED_LEGS_EXAMINED','PARENT_NOT_RESOLVED','LEGS_WITH_STORNO_PARENT','DOCUMENT_LEG_STATE_MISMATCH','ACTIVE_AMB_STORNO_PARENT','STORNO_AMB_ACTIVE_PARENT')});parent_status=Counter();parent_docs=set();troubled_docs=set()
    review=[];candidate_rows=[];case_by_ordinal={}
    unknown_parent_status=[]
    # Inspect all business-document AMB rows and all associated storno evidence.
    for ix,r in enumerate(source,1):
        doctype=f.key(r.get('DokumentTip'))
        if doctype not in f.DOC_PARENT:continue
        docid=trim(r.get('DokumentID'))
        pk_table=f.DOC_PARENT[doctype][0]
        parent=idx[pk_table].get(docid) if docid else None
        if parent is None:
            stats['PARENT_NOT_RESOLVED']+=1
            continue
        p_ordinal,p_row=parent
        ps=flag_state(p_row)
        ls=flag_state(r)
        if ps=='UNKNOWN':
            unknown_parent_status.append((ix,p_ordinal))
            continue
        parent_docs.add((doctype,docid))
        parent_status[ps]+=1
        stats['LINKED_LEGS_EXAMINED']+=1
        if ps=='STORNO':
            stats['LEGS_WITH_STORNO_PARENT']+=1
            troubled_docs.add((doctype,docid))
        if ps!=ls:
            stats['DOCUMENT_LEG_STATE_MISMATCH']+=1
            if ps=='STORNO' and ls=='ACTIVE':stats['ACTIVE_AMB_STORNO_PARENT']+=1
            if ps=='ACTIVE' and ls=='STORNO':stats['STORNO_AMB_ACTIVE_PARENT']+=1
        if ps!='STORNO' and ps==ls:continue
        group=doc_buckets[(doctype,docid)]
        same_doc_active=sum(flag_state(source[k-1])=='ACTIVE' for k in group)
        same_doc_storno=sum(flag_state(source[k-1])=='STORNO' for k in group)
        # Exact matching quantity is evidence only; no tolerance or normalization.
        qtyfield={'otkup':'KolAmbalaze','om-izlaz-koop':'KolAmbIzdata',
                  'otpremnica':'KolAmbalaze',
                  'prijemnica':'KolAmbVracena' if f.key(r.get('Smer'))=='izlaz' else 'KolAmbalaze'}[doctype]
        expected=f.dec(p_row.get(qtyfield,''));got=f.dec(r.get('Kolicina'))
        qty_status=('MATCH' if expected is not None and got is not None and expected==got else
                    'DIFFERENT' if expected is not None and got is not None else 'NOT_COMPARABLE')
        context_station=('MATCH' if f.key(r.get('EntitetTip'))=='stanica' and
                         f.key(r.get('EntitetID'))==f.key(p_row.get('StanicaID')) else 'MISMATCH') if doctype=='otpremnica' else 'NOT_APPLICABLE'
        context_driver=('MATCH' if f.key(r.get('VozacID')) and
                        f.key(r.get('VozacID'))==f.key(p_row.get('VozacID')) else 'MISMATCH') if doctype=='otpremnica' else 'NOT_APPLICABLE'
        vec=contribution_vector(r)
        inverse_legs=[k for k in vectors.get(negated_vector(vec),[]) if k!=ix] if vec else []
        inverse_legs.sort(key=lambda k:(
            (f.key(source[k-1].get('DokumentTip')),trim(source[k-1].get('DokumentID')))!=(doctype,docid),
            f.key(source[k-1].get('Datum'))!=f.key(r.get('Datum')), k))
        exact_doc_inverse=sum((f.key(source[k-1].get('DokumentTip')),trim(source[k-1].get('DokumentID')))==(doctype,docid) for k in inverse_legs)
        issue='PARENT_STORNO_AMB_ACTIVE' if ps=='STORNO' and ls=='ACTIVE' else 'PARENT_STORNO_AMB_STORNO' if ps=='STORNO' else 'PARENT_ACTIVE_AMB_STORNO'
        case={'source_row_ordinal':ix,'AmbID':r['AmbID'],'document_type':r['DokumentTip'],
              'document_id':docid,'document_parent_row_ordinal':p_ordinal,
              'amb_stornirano_raw':r.get('Stornirano',''),'parent_stornirano_raw':p_row.get('Stornirano',''),
              'amb_state':ls,'parent_state':ps,'issue_class':issue,
              'tip_ambalaze':r.get('TipAmbalaze',''),'qty_amb':r.get('Kolicina',''),
              'qty_parent':p_row.get(qtyfield,''),'quantity_check':qty_status,
              'otpremnica_station_check':context_station,'otpremnica_driver_check':context_driver,
              'same_doc_total_amb_legs':len(group),'same_doc_active_legs':same_doc_active,
              'same_doc_storno_legs':same_doc_storno,
              'exact_active_inverse_same_doc_count':exact_doc_inverse,
              'possible_active_inverse_other_doc_count':len(inverse_legs)-exact_doc_inverse,
              'interpretation':'REVIEW_ONLY_NO_AUTOMATIC_CANCELLATION'}
        review.append(case);case_by_ordinal[ix]=case
        if ls!='ACTIVE':continue
        # Preserve possible offset references locally, but never mistake their
        # equality for a storno link. Record capped detail, full counts above.
        for other_ix in inverse_legs[:30]:
            other=source[other_ix-1]
            same_doc=(f.key(other.get('DokumentTip')),trim(other.get('DokumentID')))==(doctype,docid)
            candidate_rows.append({
                'case_source_ordinal':ix,'case_amb_id':r['AmbID'],
                'candidate_source_ordinal':other_ix,'candidate_amb_id':other['AmbID'],
                'candidate_document_type':other.get('DokumentTip',''),
                'candidate_document_id':other.get('DokumentID',''),
                'same_document_id_and_type':'YES' if same_doc else 'NO',
                'same_date':'YES' if f.key(other.get('Datum'))==f.key(r.get('Datum')) else 'NO',
                'same_physical_accounts_in_inverse':'YES',
                'candidate_stornirano_raw':other.get('Stornirano',''),
                'candidate_quantity':other.get('Kolicina',''),
                'relation':'PHYSICAL_OFFSET_CANDIDATE_NOT_STORNO_PROOF'})
    if unknown_parent_status:
        raise AuditError('Unknown parent Stornirano token; example AMB ordinal '+str(unknown_parent_status[0][0]))
    if stats['LINKED_LEGS_EXAMINED']+stats['PARENT_NOT_RESOLVED'] != sum(f.key(r.get('DokumentTip')) in f.DOC_PARENT for r in source):
        raise AuditError('Parent-link accounting invariant failed')
    if stats['DOCUMENT_LEG_STATE_MISMATCH']!=stats['ACTIVE_AMB_STORNO_PARENT']+stats['STORNO_AMB_ACTIVE_PARENT']:
        raise AuditError('Status mismatch count invariant failed')
    focus=case_by_ordinal.get(focus_ordinal)
    # StornoVeze has legacy-specific semantics; scanning for literal references
    # is safe evidence discovery, not interpretation.
    storno_path=stage/'UNCONVERTED__tblStornoVeze.jsonl'
    storno_matches=[]
    if storno_path.is_file() and focus is not None:
        sv=get_rows(stage,'UNCONVERTED__tblStornoVeze')
        for n,z in enumerate(sv,1):
            if any(trim(v) in {trim(focus_row['AmbID']),trim(focus['document_id'])} for v in z.values()):
                storno_matches.append((n,z))
    focus_summary={
        'requested_source_data_ordinal':focus_ordinal,
        'focus_amb_id_found_at_ordinal':focus_id_ordinal,
        'id_and_ordinal_same_source_record':focus_same,
        'document_type_category':('Otpremnica' if f.key(focus_row.get('DokumentTip'))=='otpremnica' else 'OTHER'),
        'parent_found':focus is not None,
        'amb_state':flag_state(focus_row),
        'parent_state':focus['parent_state'] if focus else 'UNRESOLVED',
        'parent_amb_quantity_exact_match':focus['quantity_check']=='MATCH' if focus else False,
        'parent_station_matches':focus['otpremnica_station_check']=='MATCH' if focus else False,
        'parent_driver_matches':focus['otpremnica_driver_check']=='MATCH' if focus else False,
        'same_document_amb_leg_count':focus['same_doc_total_amb_legs'] if focus else 0,
        'same_document_active_amb_legs':focus['same_doc_active_legs'] if focus else 0,
        'same_document_storno_amb_legs':focus['same_doc_storno_legs'] if focus else 0,
        'same_document_exact_active_physical_offset_candidates':focus['exact_active_inverse_same_doc_count'] if focus else 0,
        'other_document_active_physical_offset_candidates':focus['possible_active_inverse_other_doc_count'] if focus else 0,
        'literal_matches_in_stornoveze':len(storno_matches),
        'approval':'NONE'
    }
    summary={'format':VERSION,'status':'STORNO_LINK_FORENSICS_ONLY_NOT_IMPORT_APPROVAL',
             'source_sha256':manifest['source_sha256'],'schema_sha256':manifest['schema_sha256'],
             'source_amb_rows':len(source),'linked_storno_document_counts':{
                 'distinct_parents_with_amb_links':len(parent_docs),
                 'distinct_storno_parents_with_amb_links':len(troubled_docs),
                 **dict(stats)},
             'focus':focus_summary,
             'review_cases':len(review),'active_offset_candidates_listed_in_private_csv':len(candidate_rows),
             'notes':['An ACTIVE AMB leg linked to a STORNO parent is a proven legacy status conflict, not a proven missing cancellation.',
                      'Equal amounts or physically opposite active legs do not establish that they cancel the same historical transaction.',
                      'Unknown source ownership, business movement kind, and chronology remain unresolved.',
                      'No source data or target XLSM mutated; no AMB-10 entries, purchases or inverses generated.',
                      'Sensitive IDs, exact amounts, StornoVeze content and candidate details are PRIVATE locally.']}
    return summary,review,candidate_rows,focus_row,focus,storno_matches,source,doc_buckets


def focus_text(focus_ordinal,source,focus,storno_matches,doc_buckets):
    row=source[focus_ordinal-1]
    lines=['AGRIX MIG-04.6 / PRIVATE / Q source ordinal '+str(focus_ordinal),
           'AmbID: '+trim(row.get('AmbID')),
           'AMB Stornirano: '+repr(row.get('Stornirano',''))+'; resolved: '+flag_state(row),
           'DokumentTIP: '+trim(row.get('DokumentTip')),
           'DokumentID: '+trim(row.get('DokumentID')),
           'Tip ambalaze: '+trim(row.get('TipAmbalaze')),
           'Kolicina: '+trim(row.get('Kolicina')),
           'Smer/entitet/vozac: '+ ' | '.join(trim(row.get(k)) for k in ('Smer','EntitetTip','EntitetID','VozacID')),
           'Datum: '+trim(row.get('Datum'))]
    if focus:
        lines+=['Roditelj redni broj: '+str(focus['document_parent_row_ordinal']),
                'Roditelj Stornirano: '+repr(focus['parent_stornirano_raw'])+'; resolved: '+focus['parent_state'],
                'Kolicina roditelja: '+trim(focus['qty_parent']),
                'Dokument/ambalaza quantity check: '+focus['quantity_check'],
                'Otpremnica station/driver context: '+focus['otpremnica_station_check']+' / '+focus['otpremnica_driver_check'],
                'Problem: '+focus['issue_class'],
                'Potencijalni aktivni kontra-potezi, ISTI dokument: '+str(focus['exact_active_inverse_same_doc_count']),
                'Potencijalni aktivni kontra-potezi, DRUGI dokument: '+str(focus['possible_active_inverse_other_doc_count'])]
    else:
        lines+=['No linked parent status review case found for this source ordinal.']
    docs=doc_buckets.get((f.key(row.get('DokumentTip')),trim(row.get('DokumentID'))),[])
    lines+=['','Svi redovi ambalaze pod ISTIM DokumentTip + DokumentID:']
    for ix in docs:
        a=source[ix-1]
        lines.append('  data ordinal '+str(ix)+'; AmbID='+trim(a.get('AmbID'))+
                     '; stanje='+flag_state(a)+'; smer='+trim(a.get('Smer'))+
                     '; kolicina='+trim(a.get('Kolicina'))+
                     '; entitet='+trim(a.get('EntitetTip'))+'/'+trim(a.get('EntitetID'))+
                     '; vozac='+trim(a.get('VozacID')))
    lines+=['','StornoVeze literal reference hits: '+str(len(storno_matches))]
    for i,z in storno_matches:
        lines.append('  StornoVeze row '+str(i)+': '+json.dumps(z,ensure_ascii=False,sort_keys=True))
    lines+=['','NONE of these checks authorize manual edits, deletion or AMB-10 posting.']
    return '\n'.join(lines)+'\n'


def render_summary(s):
    counts=s['linked_storno_document_counts'];q=s['focus']
    lines=['AGRIX MIG-04.6 / STORNO-LINEAGE READ ONLY','Source SHA256: '+s['source_sha256'],
           'AMB legacy rows: '+str(s['source_amb_rows']),
           'AMB with verified parent: '+str(counts.get('LINKED_LEGS_EXAMINED',0)),
           'Amb active / parent storno: '+str(counts.get('ACTIVE_AMB_STORNO_PARENT',0)),
           'Amb storno / parent active: '+str(counts.get('STORNO_AMB_ACTIVE_PARENT',0)),
           'Distinct storno parents with AMB references: '+str(counts['distinct_storno_parents_with_amb_links']),
           'Q focus ordinal: '+str(q['requested_source_data_ordinal'])+', matching requested AmbID ordinal: '+str(q['focus_amb_id_found_at_ordinal']),
           'Focus IDs same record: '+str(q['id_and_ordinal_same_source_record']),
           'Focus amb/parent states: '+q['amb_state']+'/'+q['parent_state'],
           'Focus quantity match: '+str(q['parent_amb_quantity_exact_match']),
           'Focus station/driver match: '+str(q['parent_station_matches'])+'/'+str(q['parent_driver_matches']),
           'Focus same-document AMB legs: '+str(q['same_document_amb_leg_count']),
           'Focus exact active opposite-physical candidates, same document: '+str(q['same_document_exact_active_physical_offset_candidates']),
           'Focus possible offsets with OTHER document: '+str(q['other_document_active_physical_offset_candidates']),
           'Focus StornoVeze literal reference hits: '+str(q['literal_matches_in_stornoveze']),
           'NO business ownership, cancellation or import authorization. No files changed.']
    return '\n'.join(lines)+'\n'


def cli(argv=None):
    p=argparse.ArgumentParser(description='Forensic Q-00700 and all storno-parent linked AMB evidence from pinned _v2 staging')
    p.add_argument('--stage',required=True,type=Path)
    p.add_argument('--out',required=True,type=Path)
    p.add_argument('--focus-ordinal',type=int,default=700)
    p.add_argument('--focus-amb-id',default='AMB-00700')
    a=p.parse_args(argv)
    stage=a.stage.resolve();out=a.out.resolve()
    if not stage.is_dir():raise AuditError('Missing MIG-02 private staging directory')
    if out==stage or stage in out.parents or out in stage.parents:raise AuditError('Input and output overlap')
    if out.exists():raise AuditError('Output folder already exists, choose a new one; do not overwrite')
    s,review,candidates,focus_row,focus,storno_matches,source,buckets=build(stage,a.focus_ordinal,a.focus_amb_id)
    out.mkdir(parents=True,exist_ok=False)
    try:
        (out/'agrix_storno_lineage_summary.json').write_text(json.dumps(s,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
        (out/'agrix_storno_lineage_summary.txt').write_text(render_summary(s),encoding='utf-8')
        f.private_csv(out/'storno_parent_amb_cases_PRIVATE.csv',review,[
            'source_row_ordinal','AmbID','document_type','document_id','document_parent_row_ordinal',
            'amb_stornirano_raw','parent_stornirano_raw','amb_state','parent_state',
            'issue_class','tip_ambalaze','qty_amb','qty_parent','quantity_check',
            'otpremnica_station_check','otpremnica_driver_check',
            'same_doc_total_amb_legs','same_doc_active_legs','same_doc_storno_legs',
            'exact_active_inverse_same_doc_count','possible_active_inverse_other_doc_count','interpretation'])
        f.private_csv(out/'storno_physical_offset_candidates_PRIVATE.csv',candidates,[
            'case_source_ordinal','case_amb_id','candidate_source_ordinal','candidate_amb_id',
            'candidate_document_type','candidate_document_id','same_document_id_and_type','same_date',
            'same_physical_accounts_in_inverse','candidate_stornirano_raw','candidate_quantity','relation'])
        focus_path=out/'Q00700_local_evidence_PRIVATE.txt'
        focus_path.write_text(focus_text(a.focus_ordinal,source,focus,storno_matches,buckets),encoding='utf-8')
        try:
            import os
            os.chmod(focus_path,0o600)
        except OSError:pass
    except BaseException:
        for x in out.iterdir():
            if x.is_file():x.unlink()
        out.rmdir()
        raise
    print(render_summary(s))
    print('Privatni podaci ostaju lokalno. Podeli samo agrix_storno_lineage_summary.json i .txt.')
    return 0

if __name__=='__main__':
    try:raise SystemExit(cli())
    except (AuditError,OSError,ValueError,KeyError,json.JSONDecodeError) as e:
        print('MIG-04.6 STOP:',e,file=sys.stderr)
        raise SystemExit(2)
