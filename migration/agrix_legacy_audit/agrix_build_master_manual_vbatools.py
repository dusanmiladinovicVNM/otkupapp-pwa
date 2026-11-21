#!/usr/bin/env python3
"""Build/verify a clean VBA bootstrap with modVbaTools and modLogo stub.

MIG-08.2 verification permits exact documented runtime audit suffixes for
tblPrevoznici, tblUtovar and tblUtovarStavke while preserving the complete
canonical ordered schema prefix on every table.

Requires Windows + Microsoft Excel + pywin32 and the pinned local Git checkout.
Never touches customer workbooks, stage data or a pre-existing output file.
The generated workbook is NOT approved until VBA re-open/export succeeds and
manual ImportAllVBA has restored canonical modLogo, Compile and re-open are verified.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

PINNED_COMMIT = '66f7897fa5231786f7d3d3dd3e566922c8e80ed0'
PINNED_SCHEMA = 'b9916598d1f8d29b8e66b3fd0805d36f4c2e695df48fe6ef32155a7b6a765d73'
TABLE_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PENDING_MODULE = 'modLogo'
TOOLS_MODULE = 'modVbaTools'
STUB_TAG = 'AGRIx_MIG081_LOGO_BOOTSTRAP'
# modSetup.EnsureRuntimeSchema appends these four audit fields to three tables.
# schema/schema.json remains the immutable, ordered prefix; no other additions
# are accepted by this pinned MIG-08 bootstrap verifier.
RUNTIME_AUDIT_SUFFIX = ('CreatedAt', 'CreatedBy', 'ModifiedAt', 'ModifiedBy')
ALLOWED_RUNTIME_SUFFIXES = {
    'tblUtovar': RUNTIME_AUDIT_SUFFIX,
    'tblUtovarStavke': RUNTIME_AUDIT_SUFFIX,
    'tblPrevoznici': RUNTIME_AUDIT_SUFFIX,
}


def canonical_table_structure(expected, actual):
    """Return verified runtime additions, or fail on any real schema drift.

    The canonical names/order must match exactly at every canonical position.
    Only the complete known post-canonical audit suffix is accepted on the
    three tables explicitly self-healed by modSetup.EnsureRuntimeSchema.
    """
    if set(actual) != set(expected):
        missing = sorted(set(expected)-set(actual))
        extra = sorted(set(actual)-set(expected))
        raise RuntimeError(f'Workbook table set mismatch missing={missing} extra={extra}')
    extensions = {}
    for key in sorted(expected):
        canonical = list(expected[key])
        observed = list(actual[key])
        if observed[:len(canonical)] != canonical:
            raise RuntimeError(f'Workbook canonical schema mismatch {key}: '
                               f'expected prefix {canonical}, got {observed}')
        tail = observed[len(canonical):]
        permitted = ALLOWED_RUNTIME_SUFFIXES.get(key)
        if tail and (permitted is None or tail != list(permitted)):
            raise RuntimeError(f'Workbook unsupported runtime suffix {key}: '
                               f'expected none or {list(permitted) if permitted else []}, got {tail}')
        if len(set(observed)) != len(observed):
            raise RuntimeError(f'Workbook duplicate column {key}')
        if tail:
            extensions[key] = tail
    return extensions

LOGO_STUB = '''Attribute VB_Name = "modLogo"
Option Explicit
' AGRIx_MIG081_LOGO_BOOTSTRAP -- intentionally replaced by ImportAllVBA.
Public Const LOGO_SPLASH As String = "SPLASH"
Public Const LOGO_KARTICA As String = "KARTICA"
Public Const LOGO_MINI As String = "MINI"
Public Const LOGO_BG_SPLASH As Long = &H12281A
Public Const LOGO_BG_KARTICA As Long = &HEEF4F7
Public Const LOGO_BG_MINI As Long = &HEEF4F7
Public Function LogoSlika(ByVal kljuc As String) As Object
    Set LogoSlika = Nothing
End Function
Public Sub LogoOtpusti()
End Sub
Public Function LogoOdnos(ByVal kljuc As String) As Single
    LogoOdnos = 1!
End Function
Public Function LogoKljuc(ByVal osnovni As String, ByVal visinaTacaka As Single) As String
    LogoKljuc = osnovni
End Function
'''



def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def repository_gate(repo):
    repo = Path(repo).resolve()
    for name in ('schema/schema.json', 'src-vba/modVbaTools.bas',
                 'src-vba/modLogo.bas', 'src-vba/ThisWorkbook.doccls',
                 'tools/make_dev_workbook.py', 'tools/make_fixture.py'):
        if not (repo / name).is_file():
            raise RuntimeError('Repo missing required file: ' + name)
    rev = subprocess.check_output(['git', '-C', str(repo), 'rev-parse', 'HEAD'], text=True).strip()
    if rev != PINNED_COMMIT:
        raise RuntimeError('Repository HEAD mismatch (refuse build): ' + rev)
    sh = digest(repo / 'schema/schema.json')
    if sh != PINNED_SCHEMA:
        raise RuntimeError('Canonical schema hash mismatch (refuse build): ' + sh)
    # No uncommitted source/schema drift, except optional release stamp.
    paths = subprocess.check_output(['git', '-C', str(repo), 'diff', '--name-only', 'HEAD', '--',
                  'schema/schema.json', 'src-vba', 'tools/make_dev_workbook.py',
                  'tools/make_fixture.py'], text=True).splitlines()
    drift = [p for p in paths if p.replace('\\','/') != 'src-vba/modBuildInfo.bas']
    if drift:
        raise RuntimeError('Uncommitted code/schema drift: ' + ', '.join(drift[:8]))
    src_files = list((repo / 'src-vba').iterdir())
    candidates = {p.stem for p in src_files if p.suffix.lower() in ('.bas','.cls','.frm')}
    if PENDING_MODULE not in candidates or 'modLogo' not in candidates:
        raise RuntimeError('Expected source module missing')
    return repo, candidates


def stage_sources(src, tempdir):
    """Bootstrap must have modVbaTools, but modLogo starts as a valid stub.

    ImportAllVBA is the ONLY step allowed to replace that stub from canonical
    repo sources. NEVER run business app while stub remains.
    """
    source = Path(tempdir) / 'src-vba'
    source.mkdir()
    copied = []
    for path in Path(src).iterdir():
        if not path.is_file():
            continue
        target = source / path.name
        if path.name.casefold() == 'modlogo.bas':
            target.write_bytes(LOGO_STUB.replace('\n', '\r\n').encode('ascii'))
        else:
            shutil.copy2(path, target)
        copied.append(path.name)
    if copied.count('modLogo.bas') != 1 or copied.count('modVbaTools.bas') != 1:
        raise RuntimeError('Bootstrap must contain exactly one logo stub and modVbaTools')
    if digest(source / 'modLogo.bas') == digest(Path(src) / 'modLogo.bas'):
        raise RuntimeError('modLogo stub has same hash as canonical source -- no delta')
    if digest(source / 'modVbaTools.bas') != digest(Path(src) / 'modVbaTools.bas'):
        raise RuntimeError('modVbaTools differs from canonical source')
    return source


def import_builder(repo):
    target = repo / 'tools/make_dev_workbook.py'
    spec=importlib.util.spec_from_file_location('agrix_canonical_builder',target)
    if spec is None or spec.loader is None:
        raise RuntimeError('Cannot load canonical generator')
    module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def workbook_structure(path, schema):
    expected={row['table']:row['columns'] for row in schema['tables']}
    with zipfile.ZipFile(path) as z:
        bad=z.testzip()
        if bad: raise RuntimeError('Corrupt OOXML entry: '+bad)
        names=[p for p in z.namelist() if p.startswith('xl/tables/table') and p.endswith('.xml')]
        actual={}
        for p in names:
            e=ET.fromstring(z.read(p))
            table_name=e.get('name')
            if not table_name or table_name in actual:
                raise RuntimeError('Missing/duplicate Excel table name')
            container=e.find(TABLE_NS+'tableColumns')
            if container is None: raise RuntimeError('No table columns: '+table_name)
            actual[table_name]=[c.get('name') for c in container.findall(TABLE_NS+'tableColumn')]
        # The app itself validates a canonical *prefix* because some
        # runtime-only columns are appended by EnsureRuntimeSchema.
        runtime_extensions = canonical_table_structure(expected, actual)
        if 'xl/vbaProject.bin' not in z.namelist():
            raise RuntimeError('Workbook has no VBA project binary')
        vba=z.read('xl/vbaProject.bin')
        if not vba.startswith(bytes.fromhex('D0CF11E0A1B11AE1')):
            raise RuntimeError('Invalid OLE VBA project binary')
    return {'tables':len(actual),'columns':sum(map(len,actual.values())),
            'canonical_columns':sum(map(len,expected.values())),
            'runtime_columns':sum(map(len,actual.values()))-sum(map(len,expected.values())),
            'runtime_extensions':runtime_extensions,
            'vba_binary_bytes':len(vba),'vba_binary_sha256':hashlib.sha256(vba).hexdigest()}


def vbide_probe(workbook, timeout, expect_canonical_logo):
    """Read-only export in dedicated Excel, then validate logo state.

    Both pre- and post-ImportAllVBA states must load without VBA corruption.
    """
    from agrix_mig0712_vba_isolation import run_one
    with tempfile.TemporaryDirectory(prefix='agrix_master_vba_probe_') as td:
        tmp=Path(td)
        state, inv=run_one('master',workbook,tmp,tmp,timeout)
        if not inv or state.get('status') != 'EXPORT_PASS':
            raise RuntimeError('VBA read-only reopen/export FAIL: '+str(state))
        modules=set(inv['components'])
        for required in ('modLogo','modVbaTools','ThisWorkbook'):
            if required not in modules:
                raise RuntimeError('VBA probe missing '+required)
        exported=tmp/'master'/'modLogo.bas'
        content=exported.read_bytes()
        tag=STUB_TAG.encode('ascii')
        if expect_canonical_logo:
            if tag in content or len(content)<35000 or b'B64_MINI2' not in content:
                raise RuntimeError('ImportAllVBA did not restore complete canonical modLogo; DO NOT USE')
        else:
            if tag not in content:
                raise RuntimeError('Expected bootstrap stub, but different logo was found')
        return {'status':state['status'], 'components':len(modules),
                'references_count':len(inv['references']),
                'logo_state':'CANONICAL_SHAPE' if expect_canonical_logo else 'BOOTSTRAP_STUB'}


def verify(workbook, repo, expected, require_tools, timeout):
    schema=json.loads((repo/'schema/schema.json').read_text(encoding='utf-8'))
    structure=workbook_structure(workbook,schema)
    export=vbide_probe(workbook,timeout,expect_canonical_logo=require_tools)
    return {'status':'VBA_REOPEN_PASS_MANUAL_COMPILE_TO_CONFIRM' if require_tools else 'BOOTSTRAP_READY_RUN_IMPORTALLVBA',
            'workbook_sha256':digest(workbook),'schema':structure,
            'vba_reopen_export':export,'missing_vba_modules':[],
            'canonical_sha':PINNED_COMMIT,'canonical_schema_sha':PINNED_SCHEMA}


def main(argv=None):
    p=argparse.ArgumentParser(description='AgriX modVbaTools-enabled bootstrap, manual ImportAllVBA required')
    p.add_argument('--repo',default=str(Path.home()/'otkupapp-pwa'))
    p.add_argument('--out',required=True,help='NEW XLSM output; must not exist for build')
    p.add_argument('--verify-only',action='store_true',help='Verify AFTER ImportAllVBA, manual Compile and save; never modify workbook')
    p.add_argument('--timeout',type=int,default=120,help='VBA reopen/export timeout, seconds')
    a=p.parse_args(argv)
    if os.name!='nt':
        p.error('Windows Excel + pywin32 required to assemble VBA XLSM')
    if not 30<=a.timeout<=600:
        p.error('Timeout must be 30..600 sec')
    repo,expected=repository_gate(a.repo)
    out=Path(a.out).resolve()
    if out.suffix.lower()!='.xlsm':p.error('Output must end .xlsm')
    if out==repo or repo in out.parents:
        p.error('Keep output outside the source repository')
    if not a.verify_only and out.exists():
        p.error('Refusing to overwrite existing XLSM: '+str(out))
    if a.verify_only and not out.is_file():
        p.error('Workbook not found: '+str(out))
    status='NO_GO'
    try:
        if not a.verify_only:
            out.parent.mkdir(parents=True,exist_ok=True)
            print('[MASTER] source pin/schema PASS',flush=True)
            with tempfile.TemporaryDirectory(prefix='agrix_master_src_') as td:
                staged=stage_sources(repo/'src-vba',td)
                print('[MASTER] importing canonical VBA with '+PENDING_MODULE+' STUB and full '+TOOLS_MODULE,flush=True)
                builder=import_builder(repo)
                builder.SRC=str(staged)
                code=builder.napravi(str(out).replace('/','\\'))
                if code not in (None,0):
                    raise RuntimeError('Canonical Excel builder returned '+str(code))
            print('[MASTER] Excel save/close returned; verifying OOXML + independent VBA reopen',flush=True)
        report=verify(out,repo,expected,a.verify_only,a.timeout)
        status=report['status']
        print('[MASTER] '+status,flush=True)
        print('[MASTER] table_count='+str(report['schema']['tables'])+' VBA_exported='+str(report['vba_reopen_export']['components']),flush=True)
    except Exception as exc:
        report={'status':'NO_GO','error_type':type(exc).__name__, 'error':str(exc),
                'template_exists':out.is_file()}
        print('[MASTER] NO-GO: '+str(exc),file=sys.stderr,flush=True)
    report.update({'format':'AGRIx-MIG-08.2-BOOTSTRAP-VERIFY','client_data_read':False,
                   'import_all_vba_manual':True,'production_approved':False,
                   'path_suppressed':True,'verification_only':a.verify_only})
    report_path=out.parent/('agrix_master_bootstrap_verify_summary.json'
                            if a.verify_only else 'agrix_master_bootstrap_build_summary.json')
    report_path.write_text(json.dumps(report,indent=2,ensure_ascii=True)+'\n',encoding='utf-8')
    print('[MASTER] report: '+str(report_path),flush=True)
    return 0 if status in ('BOOTSTRAP_READY_RUN_IMPORTALLVBA', 'VBA_REOPEN_PASS_MANUAL_COMPILE_TO_CONFIRM') else 2


if __name__=='__main__':
    raise SystemExit(main())
