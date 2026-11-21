#!/usr/bin/env python3
"""Read-only AgriX AMB10 migration payload inventory. No customer values exported.

Identify MIG-06.x private payload files using their known SHA256 from approved
preflight reports. Collect JSONL key shapes and counts, never any source data.
"""
from __future__ import annotations
import argparse, hashlib, json, sys
from collections import Counter
from pathlib import Path

KNOWNS={
 'e2f354b3af25d0c0098be2ff838e7f9b528f89999f38c1877ad8f956bb8917aa':'MIG06_5_AMB_LEDGER_ADAPTER',
 '0269ead5b5c9e0d6bad62b0bfa65c2e929d47fbf7ad0f86130c2411cb460dd82':'MIG06_5_REVERS_HEADERS',
 '2155b5ad6c192a62f7b7f0b57d9f5a7b0d8ffcaca3ebba7273aad678f89d9215':'MIG06_4_ID_RESERVATIONS',
 '757d0bf92811f653b77ce918458e2aa3c805d3f32f9969dfcf5cf9b6f60cea1d':'MIG06_3_TARGET_DRAFT',
 'f2fb513012c8263873a02611f90cdc8caabd06212e97167ed7dc82c0d4e54bf2':'MIG06_2_REVERS_HEADER_DRAFT',
 '66d6ac47e56b8f881bc1957aa5964d0a5736f747721447bcd1ed9f5846de37b1':'MIG06_2_REVERS_LINK_DRAFT',
 '69f639410c407a5bcfd53d71912ae665a2b4982':'NOT_A_REAL_HASH',
 '8a3d1c370b95f08e602216d124af12d67febb7ae3260360789bad7abca138bb5':'MIG02_LEGACY_AMB_RAW'
}
KNOWNS.pop('69f639410c407a5bcfd53d71912ae665a2b4982')
EXPECTED={'MIG06_5_AMB_LEDGER_ADAPTER':4669,'MIG06_5_REVERS_HEADERS':48,'MIG02_LEGACY_AMB_RAW':7815}
FIELDS=('VrstaKretanja','DokumentTIP','OdNalogTip','NaNalogTip','Vrsta','Smer','EntitetTip')
SENSITIVE=('ID','Name','Ime','Naziv','Broj','Partner','Kolicina','Datum','Cena','Saldo','Iznos','Napomena')

def digest(p):
 h=hashlib.sha256()
 with p.open('rb') as f:
  for b in iter(lambda:f.read(1<<20),b''):h.update(b)
 return h.hexdigest()

def shape(p):
 suffix=p.suffix.lower()
 keys=Counter(); cat={k:Counter() for k in FIELDS}; rows=0; typ=Counter(); nesting=Counter()
 def collect(v):
  nonlocal rows
  if not isinstance(v,dict): typ[type(v).__name__]+=1;return
  rows+=1
  for k,x in v.items():
   keys[k]+=1
   if isinstance(x,dict):nesting[k]+=1
   if k in cat and x is not None:
    val=str(x).strip()
    if len(val)<=60 and all(c.isupper() or c.isdigit() or c in '_- ' for c in val):cat[k][val]+=1
    else:cat[k]['[unpublished]']+=1
 if suffix=='.jsonl':
  with p.open('r',encoding='utf-8-sig') as f:
   for line in f:
    if line.strip():collect(json.loads(line))
 elif suffix=='.json':
  data=json.loads(p.read_text('utf-8-sig'))
  if isinstance(data,list):
   for item in data:collect(item)
  elif isinstance(data,dict):
   # Never include object values; summarize top-level shape only.
   collect(data)
 else:return None
 return {'records':rows,'field_names':sorted(keys),'nested_fields':sorted(nesting),
         'enum_counts':{k:dict(v) for k,v in cat.items() if v},'non_object_types':dict(typ)}

def main():
 a=argparse.ArgumentParser()
 a.add_argument('--project-dir',type=Path,required=True)
 a.add_argument('--out',type=Path,required=True)
 a=a.parse_args();root=a.project_dir.resolve();out=a.out.resolve()
 if not root.is_dir():raise ValueError('Project directory missing')
 files=[];dirs=sorted(x for x in root.iterdir() if x.is_dir() and x.name.startswith('MigrationAmb10'))
 stage=root/'MigrationStage_v2'
 if stage.exists():dirs.append(stage)
 if not dirs:raise ValueError('No MigrationAmb10* directories found')
 for d in dirs:
  for p in d.rglob('*'):
   if p.suffix.lower() not in ('.jsonl','.json') or not p.is_file():continue
   if p.name.endswith('_summary.json'):continue
   if p.stat().st_size>200_000_000:continue
   h=digest(p);role=KNOWNS.get(h,'unmatched')
   # Unknown JSON reports are only relevant if they look like private payload.
   if role=='unmatched' and p.suffix.lower()=='.json' and p.stat().st_size<20000:continue
   if role=='unmatched' and p.suffix.lower()=='.jsonl' and p.stat().st_size<1000:continue
   try:record=shape(p)
   except (UnicodeError,ValueError,json.JSONDecodeError) as e:
    record={'parse_error':type(e).__name__}
   # No row values, identifiers, monetary amounts or account balances in this output.
   files.append({'path_relative_to_project':p.relative_to(root).as_posix(),
                 'role':role,'sha256':h,'file_bytes':p.stat().st_size,'shape':record})
 matched={v for v in (x['role'] for x in files) if v!='unmatched'}
 issues=[]
 for req in ('MIG06_5_AMB_LEDGER_ADAPTER','MIG06_5_REVERS_HEADERS'):
  cand=[x for x in files if x['role']==req]
  if len(cand)!=1:issues.append(f'{req}: expected exactly one payload, found {len(cand)}')
  elif cand[0]['shape'].get('records')!=EXPECTED[req]:issues.append(f'{req}: unexpected row count')
 report={'format':'AGRIx-MIG15-PRIVATE-PAYLOAD-SCHEMA-INVENTORY-v1','source_workbook_or_stage_modified':False,
   'target_xlsm_modified':False,'business_amounts_and_private_row_values_excluded':True,
   'directories_scanned':[x.name for x in dirs],
   'known_payloads_detected':sorted(matched),
   'matched_payload_count':sum(x['role']!='unmatched' for x in files),
   'issues':issues,'files':files}
 if out.exists():raise ValueError('Output file already exists; choose a different --out')
 out.parent.mkdir(parents=True,exist_ok=True)
 out.write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 print('Known payloads:',', '.join(sorted(matched)))
 print('Issues:',issues if issues else 'none')
 print('Summary:',out)
 if issues:return 2
 return 0
if __name__=='__main__':
 try:sys.exit(main())
 except (OSError,ValueError) as exc:
  print('PREFLIGHT BLOCKED:',type(exc).__name__,str(exc),file=sys.stderr)
  sys.exit(2)
