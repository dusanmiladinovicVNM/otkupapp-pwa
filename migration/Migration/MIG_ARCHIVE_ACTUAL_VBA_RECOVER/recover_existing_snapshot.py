#!/usr/bin/env python3
"""Recover already-exported VBA modules after live Excel locked workbook SHA read.
READ-ONLY with respect to source XLSM, source snapshots, Excel, and VBA.
"""
import argparse
import datetime as dt
import hashlib
import json
import sys
import zipfile
from pathlib import Path

MODULES = ('modAmbalaza', 'modIzvestaj', 'modScrIzvestaji', 'modAmbalazaUgovor')


def sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def gather(snapshot: Path):
    parts = {}
    manifest = []
    for name in MODULES:
        x = {}
        for ext in ('bas', 'txt'):
            p = snapshot / f'{name}.{ext}'
            if not p.is_file() or p.stat().st_size < 50:
                raise ValueError(f'Missing/empty module: {p}')
            b = p.read_bytes()
            x[ext] = b
            parts[p.name] = b
        decoded = x['txt'].decode('utf-8-sig', errors='strict')
        if len(decoded.splitlines()) < 20:
            raise ValueError(f'Unexpectedly short source: {name}')
        manifest.append({
            'module': name,
            'source_lines': len(decoded.splitlines()),
            'bas_sha256': sha256(x['bas']),
            'txt_sha256': sha256(x['txt']),
            'bas_bytes': len(x['bas']),
            'txt_bytes': len(x['txt'])
        })
    return parts, manifest


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--base', required=True, help='Folder ACTUAL_VBA_EXPORT next to the archived XLSM')
    a = ap.parse_args()
    base = Path(a.base).resolve()
    if base.name.lower() != 'actual_vba_export' or base.parent.name.lower() != 'migrationarchivexlsm_05':
        raise SystemExit('STOP: expected .../MigrationArchiveXLSM_05/ACTUAL_VBA_EXPORT')
    if not base.is_dir():
        raise SystemExit(f'STOP: folder does not exist: {base}')
    snapshots = sorted((p for p in base.iterdir() if p.is_dir() and p.name.startswith('snapshot_')),
                       key=lambda p: p.stat().st_mtime, reverse=True)
    if not snapshots:
        raise SystemExit('STOP: no previously exported snapshot_* folder; use the original exporter first')
    candidates = []
    for s in snapshots:
        try:
            parts, manifest = gather(s)
            candidates.append((s, parts, manifest))
            break
        except (ValueError, UnicodeError, OSError) as exc:
            print(f'SKIP_INCOMPLETE {s}: {exc}', file=sys.stderr)
    if not candidates:
        raise SystemExit('STOP: no complete snapshot with all 4 .bas and 4 .txt files')
    snapshot, parts, manifest = candidates[0]
    tag = dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    zip_path = base / f'ACTUAL_VBA_SNAPSHOT_RECOVERED_{tag}.zip'
    details = {
        'status': 'ACTUAL_VBA_SNAPSHOT_RECOVERED',
        'source_snapshot_folder': str(snapshot),
        'modules': manifest,
        'workbook_sha256': None,
        'workbook_sha256_note': 'Not attempted: Excel has locked the XLSM for direct hashing',
        'original_workbook_read': False,
        'original_workbook_modified': False,
        'excel_com_used': False,
        'vba_modified': False,
        'snapshot_sources_modified': False,
        'created_at': dt.datetime.now(dt.timezone.utc).isoformat(),
    }
    with zipfile.ZipFile(zip_path, 'x', compression=zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
        z.writestr('SOURCE_MANIFEST.json', json.dumps(details, ensure_ascii=False, indent=2))
    with zipfile.ZipFile(zip_path) as z:
        err = z.testzip()
        if err or len(z.namelist()) != 9:
            raise RuntimeError(f'ZIP validation failed: {err}')
    print('ACTUAL_VBA_SNAPSHOT_RECOVERED')
    for m in manifest:
        print(f"EXPORTED {m['module']} ({m['source_lines']} lines)")
    print(f'SNAPSHOT_SOURCE={snapshot}')
    print(f'SNAPSHOT_ZIP={zip_path}')
    print('NO_EXCEL_OPENED_NO_XLSM_MODIFIED')


if __name__ == '__main__':
    main()
