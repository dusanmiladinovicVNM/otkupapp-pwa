#!/usr/bin/env python3
"""MIG-10.2: propose/apply narrow, gated VBA read policy for a legacy cancelled zero-kg otkup.

No customer workbook access or data mutation. Local private staging is inspected only
for provenance and exact VBA-compatible Stornirano='Da' before proposing code changes.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import subprocess
import sys
import difflib
from pathlib import Path
from decimal import Decimal

from agrix_mig10_otkup_contract_audit import audit, load_jsonl, EXPECTED, AuditError, SCHEMA_SHA

EXPECTED_COMMIT = '66f7897fa5231786f7d3d3dd3e566922c8e80ed0'
ANCHORS = [
 ("    RequireJedinstvenZaglavljeOtkupa zagl, SRC\n\n    Dim imaStavku As Object\n",
  "    RequireJedinstvenZaglavljeOtkupa zagl, SRC\n\n"
  "    ' MIG-10.2: historical cancelled documents may retain a zero-kg line.\n"
  "    ' Live/new documents are STILL strictly positive; no source rows are changed.\n"
  "    Dim storniranaZaglavlja As Object\n"
  "    Set storniranaZaglavlja = OtkupStorniranaZaglavlja()\n\n"
  "    Dim imaStavku As Object\n"),
 ("            RequireStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _\n",
  "            RequireStavkaUgovor zagl, storniranaZaglavlja, oid, d(i, cKl), d(i, cKol), d(i, cCena), _\n"),
 ("Private Sub RequireStavkaUgovor(ByVal zagl As Object, ByVal oid As String, _\n",
  "Private Sub RequireStavkaUgovor(ByVal zagl As Object, ByVal storniranaZaglavlja As Object, _\n"
  "                                ByVal oid As String, _\n"),
 ("    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then\n        Err.Raise vbObjectError + 1906, sourceName, _\n",
  "    ' Only ZERO kg of an already CANCELLED header is a readable historic exception.\n"
  "    ' Negative kg, zero/negative price and active zero kg remain hard failures.\n"
  "    If CDbl(kol) < 0 Or CDbl(cena) <= 0 Or _\n"
  "       (CDbl(kol) = 0 And Not storniranaZaglavlja.Exists(oid)) Then\n"
  "        Err.Raise vbObjectError + 1906, sourceName, _\n"),
]
HELPER = """' MIG-10.2: a cancelled historical line with 0 kg still exists as a document.
' This is used ONLY by the historical bulk reader. VrednostOtkupa and the new
' transaction writer keep their strict >0 contract; zero kg cannot create a new
' valid document or be settled as a positive-quantity purchase.
Private Function OtkupStorniranaZaglavlja() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    Dim d As Variant
    d = GetTableData(TBL_OTKUP)
    If IsArray(d) Then
        Dim cId As Long, cSt As Long
        cId = RequireColumnIndex(TBL_OTKUP, COL_OTK_ID, "OtkupStorniranaZaglavlja")
        cSt = RequireColumnIndex(TBL_OTKUP, COL_STORNIRANO, "OtkupStorniranaZaglavlja")
        Dim i As Long, oid As String
        For i = 1 To UBound(d, 1)
            oid = Trim$(NzToText(d(i, cId)))
            If Len(oid) > 0 Then
                If UCase$(Trim$(NzToText(d(i, cSt)))) = "DA" Then result(oid) = True
            End If
        Next i
    End If

    Set OtkupStorniranaZaglavlja = result
End Function

"""
INSERT_BEFORE = "' JEDAN LOGICKI DOKUMENT = JEDAN ID (A1/S4). Dva zaglavlja sa istim OtkupID\n"


def replace_once(body: str, old: str, new: str) -> str:
    n = body.count(old)
    if n != 1:
        raise AuditError(f'Expected exactly one canonical source anchor, got {n}: {old[:88]!r}')
    return body.replace(old, new, 1)


def patch_module(body: str) -> str:
    if 'MIG-10.2:' in body:
        raise AuditError('MIG-10.2 already present; refuse double-patching')
    original_value_function = body.split('Public Function VrednostOtkupa(', 1)[1].split('End Function', 1)[0]
    for old, new in ANCHORS:
        body = replace_once(body, old, new)
    body = replace_once(body, INSERT_BEFORE, HELPER + INSERT_BEFORE)
    after_value_function = body.split('Public Function VrednostOtkupa(', 1)[1].split('End Function', 1)[0]
    if original_value_function != after_value_function:
        raise AuditError('VrednostOtkupa unexpected mutation')
    if 'CreateOtkup_TX' not in body:
        raise AuditError('Canonical writer missing')
    if body.count('OtkupStorniranaZaglavlja()') != 2:
        raise AuditError('Unexpected helper references')
    return body


def git(repo: Path, *args: str) -> str:
    out = subprocess.run(['git', '-C', str(repo), *args], text=True, capture_output=True)
    if out.returncode:
        raise AuditError('git check failed: ' + (out.stderr or out.stdout).strip()[:400])
    return out.stdout.strip()


def precheck(repo: Path, stage: Path, ordinal: int) -> dict:
    head = git(repo, 'rev-parse', 'HEAD')
    if head != EXPECTED_COMMIT:
        raise AuditError('HEAD does not match pinned migration commit; do not patch unknown source')
    schema_path = repo/'schema'/'schema.json'
    if hashlib.sha256(schema_path.read_bytes()).hexdigest() != SCHEMA_SHA:
        raise AuditError('schema changed since MIG-09')
    previous = audit(stage)
    if previous['issue_counts'] != {'QUANTITY_NONPOSITIVE': 1} or previous['affected_source_row_count'] != 1:
        raise AuditError('MIG-10 source evidence changed')
    entries = previous['issues_by_source_row_ordinal']
    if len(entries) != 1 or entries[0]['source_row_ordinal'] != ordinal:
        raise AuditError('Unexpected MIG-10 ordinal')
    src, head, line = (load_jsonl(stage/(f+'.jsonl'),EXPECTED)[ordinal-1]
                       for f in ('SOURCE__tblOtkup','tblOtkup','tblOtkupStavke'))
    oid = str(src.get('OtkupID','')).strip()
    if not oid or oid != str(head.get('OtkupID','')).strip() or oid != str(line.get('OtkupID','')).strip():
        raise AuditError('Source/header/line identity mismatch')
    stsrc = str(src.get('Stornirano','')).strip().upper()
    sthead = str(head.get('Stornirano','')).strip().upper()
    if stsrc != 'DA' or sthead != 'DA':
        # UI filter tests exactly UCase$(Stornirano) = 'DA': do not weaken to generic TRUE/1.
        raise AuditError('Cancelled row does not have exact UI-native Stornirano=Da; migration status normalization must be audited first')
    if (Decimal(str(src.get('Kolicina','')).strip()) != 0 or
        Decimal(str(line.get('Kolicina','')).strip()) != 0 or
        Decimal(str(line.get('Cena','')).strip()) <= 0 or
        Decimal(str(line.get('KolAmbalaze','')).strip()) <= 0):
        raise AuditError('Legacy zero-kg/positive-price/positive-packaging evidence changed')
    return {'stage_counts': previous['stage_rows'], 'exception_ordinal': ordinal, 'status_exact_ui_da': True,
            'historical_rows_unchanged': True, 'quantity_class': 'ZERO', 'price_class': 'POSITIVE',
            'packaging_class': 'POSITIVE'}


def cli() -> int:
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', required=True)
    parser.add_argument('--stage', required=True)
    parser.add_argument('--out', required=True)
    parser.add_argument('--ordinal', type=int, default=123)
    parser.add_argument('--apply', action='store_true', help='Opt-in: modify only src-vba/modOtkup.bas; requires clean feature branch')
    a=parser.parse_args()
    repo=Path(a.repo).expanduser().resolve()
    stage=Path(a.stage).expanduser().resolve()
    out=Path(a.out).expanduser().resolve()
    if out.exists(): raise AuditError('Choose a fresh output directory')
    if stage == out or stage in out.parents or out in stage.parents:
        raise AuditError('Output must be outside private stage')
    evidence=precheck(repo,stage,a.ordinal)
    p=repo/'src-vba'/'modOtkup.bas'
    raw=p.read_bytes()
    nl='\r\n' if b'\r\n' in raw else '\n'
    code=raw.decode('utf-8').replace('\r\n','\n')
    new=patch_module(code)
    newbytes=new.replace('\n',nl).encode('utf-8')
    if a.apply:
        branch=git(repo,'branch','--show-current')
        if branch in ('main','master',''):
            raise AuditError('Refuse editing main: git switch -c migration/mig102-otkup-history first')
        if git(repo,'status','--porcelain'):
            raise AuditError('Working tree is not clean: no automatic edit')
        p.write_bytes(newbytes)
    diff=''.join(difflib.unified_diff(code.splitlines(keepends=True),new.splitlines(keepends=True),
                        fromfile='a/src-vba/modOtkup.bas',tofile='b/src-vba/modOtkup.bas'))
    if not diff: raise AuditError('Empty patch')
    out.mkdir(parents=True,exist_ok=False)
    (out/'MIG102_modOtkup_proposed.patch').write_text(diff,encoding='utf-8')
    rep={'format':'MIG-10.2-CANCELLED-ZERO-OTKUP-READ-POLICY','status':'PATCH_APPLIED' if a.apply else 'PROPOSAL_ONLY',
         'old_source_sha256':hashlib.sha256(raw).hexdigest(),
         'proposed_source_sha256':hashlib.sha256(newbytes).hexdigest(),
         'evidence':evidence,'customer_workbooks_opened':False,'customer_workbooks_modified':False,
         'stage_modified':False,'repo_modified':a.apply,
         'vba_compile_executed':False,'production_approved':False,
         'guard':'bulk reader only: zero kg allowed iff parent Stornirano=Da and price>0; active/negative/zero-price fail',
         'VrednostOtkupa_unchanged':True,'CREATE_writer_unchanged':True,
         'warning':'Other consumers of StavkeOtkupaRedovi require follow-up ledger/report smoke tests; do not ship before testing.'}
    (out/'agrix_mig102_summary.json').write_text(json.dumps(rep,indent=2)+'\n',encoding='utf-8')
    (out/'agrix_mig102_summary.txt').write_text('\n'.join([
        'MIG-10.2 CANCELED LEGACY ZERO: '+rep['status'],
        'Stornirano = Da (UI-compatible): VERIFIED',
        'Zero quantity with positive price and packaging: VERIFIED',
        'Source workbook and migrated XLSM: NEVER TOUCHED',
        'Only bulk historical reader adjusted; VrednostOtkupa and new writers unchanged',
        'VBA Compile and end-to-end smoke: PENDING; PRODUCTION NO-GO'])+'\n',encoding='utf-8')
    print('MIG-10.2',rep['status'],flush=True)
    print('Staging evidence: PASS, quantity ZERO, cancelled UI status DA',flush=True)
    print('Source sha',rep['old_source_sha256'],flush=True)
    print('Patch sha',rep['proposed_source_sha256'],flush=True)
    print('Local summary:',out/'agrix_mig102_summary.txt',flush=True)
    return 0

if __name__=='__main__':
    try:
        sys.exit(cli())
    except (AuditError,ValueError,OSError,UnicodeDecodeError,KeyError,json.JSONDecodeError) as e:
        print('MIG-10.2 STOP:',e,file=sys.stderr)
        sys.exit(2)
