#!/usr/bin/env python3
"""MIG-07.12 read-only VBA triage. NEVER executes macros or modifies source XLSMs.

Requires Windows Excel and pywin32. Opens one workbook PER isolated Excel process;
parent imposes a timeout and kills only the newly-created dedicated Excel PID.
All VBA component exports live in a parent-owned disposable private temp folder.
Only an aggregate, code/data-free report is written.
"""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import zipfile

FORMAT = "AGRIx-MIG-07.12-VBA-ISOLATION"
CFB_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
EXT = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}


def sha(data):
    return hashlib.sha256(data).hexdigest()


def zip_vba(path):
    with zipfile.ZipFile(path) as z:
        bad=z.testzip()
        if bad:
            raise RuntimeError('OOXML ZIP has a damaged entry: '+str(bad))
        blob=z.read('xl/vbaProject.bin')
    if not blob.startswith(CFB_SIGNATURE):
        raise RuntimeError('VBA bin does not contain an OLE compound-file signature')
    return {"vba_binary_sha256":sha(blob),"vba_binary_bytes":len(blob)}


def worker(args):
    if os.name != 'nt':
        raise RuntimeError('Windows Excel required')
    import pythoncom
    import win32com.client
    import win32process
    pythoncom.CoInitialize()
    xl=None
    wb=None
    try:
        xl=win32com.client.DispatchEx('Excel.Application')
        # Identify precisely the isolated EXCEL.EXE spawned by DispatchEx.
        hwnd=int(xl.Hwnd)
        _tid,pid=win32process.GetWindowThreadProcessId(hwnd)
        marker=Path(args.marker)
        marker.write_text(str(pid),encoding='ascii')
        xl.Visible=False
        xl.EnableEvents=False
        xl.DisplayAlerts=False
        xl.AskToUpdateLinks=False
        xl.AutomationSecurity=3  # msoAutomationSecurityForceDisable
        print('WORKER opening '+args.label+' READ ONLY',flush=True)
        wb=xl.Workbooks.Open(str(Path(args.workbook).resolve()),UpdateLinks=0,
            ReadOnly=True,IgnoreReadOnlyRecommended=True,AddToMru=False)
        if not bool(wb.ReadOnly):
            raise RuntimeError('Workbook did not open read-only')
        p=wb.VBProject
        folder=Path(args.temp)/args.label
        folder.mkdir(parents=True,exist_ok=False)
        inventory={}
        n=0
        for component in p.VBComponents:
            name=str(component.Name)
            kind=int(component.Type)
            if kind not in EXT or not name or '\\' in name or '/' in name or ':' in name:
                raise RuntimeError('Unsupported VBA component: '+name)
            if name.casefold() in {x.casefold() for x in inventory}:
                raise RuntimeError('Duplicate VBA component '+name)
            newpath=folder/(name+EXT[kind])
            print('WORKER '+args.label+' exporting '+str(n+1)+' '+name,flush=True)
            component.Export(str(newpath))
            if not newpath.is_file():
                raise RuntimeError('Export did not produce expected component file: '+name)
            exports={EXT[kind]:sha(newpath.read_bytes())}
            frx=folder/(name+'.frx')
            if kind==3 and frx.exists():
                exports['.frx']=sha(frx.read_bytes())
            inventory[name]={"kind":kind,"exports":exports}
            n+=1
        refs=sorted((str(x.Name),str(x.GUID),int(x.Major),int(x.Minor)) for x in p.References)
        if not inventory:raise RuntimeError('Empty VBA project')
        Path(args.manifest).write_text(json.dumps({"components":inventory,"references":refs},indent=2),encoding='utf-8')
        print('WORKER '+args.label+' EXPORTED '+str(n),flush=True)
        return 0
    finally:
        if wb is not None:
            try:wb.Close(SaveChanges=False)
            except Exception:pass
        if xl is not None:
            try:xl.Quit()
            except Exception:pass
        pythoncom.CoUninitialize()


def kill_dedicated_excel(marker):
    if not marker.is_file():return
    try:
        pid=int(marker.read_text(encoding='ascii').strip())
        if pid>0:
            subprocess.run(['taskkill','/PID',str(pid),'/F'],stdout=subprocess.DEVNULL,
               stderr=subprocess.DEVNULL,timeout=10,check=False)
    except (OSError,ValueError,subprocess.TimeoutExpired):pass


def run_one(label,path,root,tempdir,timeout):
    marker=tempdir/(label+'.pid')
    manifest=tempdir/(label+'.json')
    cmd=[sys.executable,str(Path(__file__).resolve()),'--worker',
         '--label',label,'--workbook',str(path),'--marker',str(marker),
         '--manifest',str(manifest),'--temp',str(tempdir)]
    print('MIG-07.12 '+label+': isolated Excel/VBIDE export starting (timeout '+str(timeout)+'s)',flush=True)
    try:
        p=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                         text=True,encoding='utf-8',errors='replace',timeout=timeout,check=False)
    except subprocess.TimeoutExpired as e:
        kill_dedicated_excel(marker)
        log=(e.stdout or b'')
        if isinstance(log,bytes):log=log.decode('utf-8','replace')
        lines=log.strip().splitlines()
        print('MIG-07.12 '+label+': TIMEOUT; last step: '+(lines[-1] if lines else 'none'),flush=True)
        return {'status':'TIMEOUT','last_step':lines[-1][-180:] if lines else 'unknown'},None
    if p.returncode!=0:
        lines=p.stdout.strip().splitlines()
        print('MIG-07.12 '+label+': FAIL ('+str(p.returncode)+'); '+(lines[-1] if lines else 'no details'),flush=True)
        # On failure, worker might not be able to close VBIDE modal; best effort cleanup
        kill_dedicated_excel(marker)
        return {'status':'FAIL','last_step':lines[-1][-280:] if lines else 'unknown'},None
    if not manifest.exists():
        return {'status':'FAIL','last_step':'Worker exited without inventory'},None
    data=json.loads(manifest.read_text(encoding='utf-8'))
    print('MIG-07.12 '+label+': EXPORT_PASS components='+str(len(data['components'])),flush=True)
    return {'status':'EXPORT_PASS','components':len(data['components'])},data


def main(argv=None):
    pa=argparse.ArgumentParser(description='Read-only, isolated VBA project comparison for MIG-07.11')
    pa.add_argument('--template',type=Path)
    pa.add_argument('--scratch',type=Path)
    pa.add_argument('--out',type=Path)
    pa.add_argument('--timeout',type=int,default=120)
    pa.add_argument('--worker',action='store_true',help=argparse.SUPPRESS)
    pa.add_argument('--label',help=argparse.SUPPRESS)
    pa.add_argument('--workbook',help=argparse.SUPPRESS)
    pa.add_argument('--marker',help=argparse.SUPPRESS)
    pa.add_argument('--manifest',help=argparse.SUPPRESS)
    pa.add_argument('--temp',help=argparse.SUPPRESS)
    a=pa.parse_args(argv)
    if a.worker:return worker(a)
    if os.name!='nt':
        raise RuntimeError('The Excel/VBIDE audit must run on Windows')
    if not a.template or not a.scratch or not a.out:
        pa.error('--template, --scratch and --out are required')
    if a.timeout<20 or a.timeout>600:pa.error('Use --timeout 20..600 seconds')
    paths={'template':a.template.resolve(),'scratch':a.scratch.resolve()}
    if not all(x.is_file() for x in paths.values()) or len(set(paths.values()))!=2:
        raise RuntimeError('Expected two distinct existing XLSM files')
    out=a.out.resolve()
    if out.exists() or any(out==p or out in p.parents or p in out.parents for p in paths.values()):
        raise RuntimeError('Output must be a NEW isolated directory not overlapping inputs')
    facts={k:zip_vba(p) for k,p in paths.items()}
    summary={'format':FORMAT,'source_files_never_saved':True,'macros_never_executed':True,
        'private_customer_data_not_exported':True,'no_production_approval':True,
        'vba_binaries':facts,'vba_raw_bytes_identical':facts['template']['vba_binary_sha256']==facts['scratch']['vba_binary_sha256'],
        'isolation':{}}
    with tempfile.TemporaryDirectory(prefix='agrix_mig0712_') as tmp:
        tempdir=Path(tmp)
        inventories={}
        for label,path in paths.items():
            status,inv=run_one(label,path,out,tempdir,a.timeout)
            summary['isolation'][label]=status
            inventories[label]=inv
        if all(x for x in inventories.values()):
            left,right=inventories['template'],inventories['scratch']
            keys=sorted(set(left['components'])|set(right['components']))
            changed=[k for k in keys if left['components'].get(k)!=right['components'].get(k)]
            refs_match=left['references']==right['references']
            summary['comparison']={'vba_component_exports_identical':not changed,
                'vba_references_identical':refs_match,'changed_component_names':changed,
                'status':'SOURCE_FORM_REFS_MATCH' if not changed and refs_match else 'VBA_DRIFT'}
        else:
            summary['comparison']={'status':'NOT_PROVEN','reason':'At least one isolated VBA load/export failed'}
    out.mkdir(parents=True,exist_ok=False)
    (out/'agrix_vba_isolation_summary.json').write_text(json.dumps(summary,indent=2)+'\n',encoding='utf-8')
    txt=(FORMAT+'\n'+'VBA raw bin SHA equality: '+str(summary['vba_raw_bytes_identical'])+'\n'
        + '\n'.join(k+': '+v['status'] for k,v in summary['isolation'].items())+'\n'
        +'Comparison: '+summary['comparison']['status']+'\n'
        +'DATA ROUNDTRIP PREVIOUSLY PASSED; BUSINESS / PRODUCTION NO-GO.\n')
    (out/'agrix_vba_isolation_summary.txt').write_text(txt,encoding='utf-8')
    print(txt,flush=True)
    print('Share ONLY agrix_vba_isolation_summary.json and .txt; never share XLSM.',flush=True)
    return 0 if summary['comparison']['status']=='SOURCE_FORM_REFS_MATCH' else 2


if __name__=='__main__':
    try:raise SystemExit(main())
    except Exception as e:
        print('MIG-07.12 STOP:',str(e),file=sys.stderr)
        raise SystemExit(2)
