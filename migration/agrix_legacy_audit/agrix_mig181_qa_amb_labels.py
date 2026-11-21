#!/usr/bin/env python3
"""MIG-18.1 QA-only: label existing ledger movements in the existing MIG-18 report.

Inputs: original 66f7897f modAmbalaza from local Git; MIG-18 generated
modIzvestaj.bas from prior QA package. Does not edit Git or workbook.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import subprocess
from pathlib import Path

BASE_REF = '66f7897fa5231786f7d3d3dd3e566922c8e80ed0'


def replace_once(text: str, old: str, new: str, name: str) -> str:
    n = text.count(old)
    if n != 1:
        raise ValueError(f'{name}: expected exactly one anchor; got {n}. No output generated.')
    return text.replace(old, new, 1)


def in_function(text: str, start_token: str, transform):
    if text.count(start_token) != 1:
        raise ValueError(f'Function anchor not unique: {start_token}')
    ix = text.index(start_token)
    end = text.index('\nEnd Function', ix) + len('\nEnd Function')
    before, body, after = text[:ix], text[ix:end], text[end:]
    changed = transform(body)
    return before + changed + after


def patch_ambalaza(s: str) -> str:
    # Contract comments precede the Public Function token, so patch before
    # slicing the function body.
    s = replace_once(s,
        "' Vraca 2D niz (1 To n, 1 To 8):",
        "' Vraca 2D niz (1 To n, 1 To 10):",
        'array contract comment')
    s = replace_once(s,
        "'   7 = ProtivpartnerTip   8 = ProtivpartnerID",
        "'   7 = ProtivpartnerTip   8 = ProtivpartnerID\n"
        "'   9 = VrstaKretanja   10 = StornoOd (prazno za original)\n"
        "' Prvih OSAM kolona ostaje nepromenjeno radi kompatibilnosti.",
        'extended array contract')
    def change(fn: str) -> str:
        fn = replace_once(fn,
            '    Dim cStorno As Long, cAmbID As Long\n',
            '    Dim cStorno As Long, cAmbID As Long, cVrstaKretanja As Long\n',
            'new movement-kind index var')
        fn = replace_once(fn,
            '    cAmbID = kolIdx(COL_AMB_ID)\n',
            '    cAmbID = kolIdx(COL_AMB_ID)\n'
            '    cVrstaKretanja = kolIdx(COL_AMB_VRSTA_KRETANJA)\n',
            'read movement kind index')
        fn = replace_once(fn,
            '    ReDim res(1 To poz.count, 1 To 8)\n',
            '    ReDim res(1 To poz.count, 1 To 10)\n',
            'append two result columns')
        fn = replace_once(fn,
            '        res(n, 6) = otkazani.Exists(AmbText(data(ri, cAmbID)))\n',
            '        res(n, 6) = otkazani.Exists(AmbText(data(ri, cAmbID)))\n'
            '        res(n, 9) = AmbText(data(ri, cVrstaKretanja))\n'
            '        res(n, 10) = AmbText(data(ri, cStorno))\n',
            'append kind and reversal source')
        return fn
    return in_function(s, 'Public Function AmbKretanjaNaloga(', change)


LABEL_FUNCTION = '''
' Prikaz vrste iz kanonske knjige, ne vrste roditeljskog dokumenta.
' VAZNO: labela nije poslovno pravilo salda, ne utice na pisca/storno.
Private Function AmbKarticaNazivKretanja(ByVal vrsta As String) As String
    Select Case UCase$(Trim$(vrsta))
        Case AMB_VK_IZDATA_PRAZNA:    AmbKarticaNazivKretanja = "izdate prazne"
        Case AMB_VK_UZ_ROBU:         AmbKarticaNazivKretanja = "primljene uz otkup"
        Case AMB_VK_POVRAT_PRAZNE:   AmbKarticaNazivKretanja = "povrat praznih"
        Case AMB_VK_PRENOS_INTERNO:  AmbKarticaNazivKretanja = "interni prenos"
        Case AMB_VK_ULAZ_TUDJE:     AmbKarticaNazivKretanja = "evidentirane tudje"
        Case AMB_VK_VRACANJE_TUDJE: AmbKarticaNazivKretanja = "vracene tudje"
        Case AMB_VK_NABAVKA:        AmbKarticaNazivKretanja = "nabavka"
        Case AMB_VK_OTPIS:          AmbKarticaNazivKretanja = "otpis"
        Case Else:                  AmbKarticaNazivKretanja = Trim$(vrsta)
    End Select
End Function
'''


def patch_izvestaj(s: str) -> str:
    # Reject base module: the running balance ordering repair must already exist.
    if "AmbKarticaSortKey(" not in s or "arr = SortArray(tmpArr, 6, True)" not in s:
        raise ValueError('modIzvestaj.bas is NOT the MIG-18 repaired version; aborting.')
    if 'Private Function AmbKarticaNazivKretanja(' in s:
        raise ValueError('Movement labels already installed; do not patch twice.')

    def change(fn: str) -> str:
        old = '''    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))
    If Not IsArray(ambData) Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If'''
        new = '''    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))
    If Not IsArray(ambData) Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If
    ' Fail closed if reader was not upgraded alongside this report.
    If UBound(ambData, 2) < 10 Then
        Err.Raise vbObjectError + 4492, SRC, _
                  "AmbKretanjaNaloga mora da vrati VrstaKretanja i StornoOd."
    End If'''
        fn = replace_once(fn,old,new,'reader contract guard')
        old = '''                    Dim opis As String
                    opis = NzToText(ambData(i, colTip))   ' TipAmbalaze
                    If lbl <> "" Then opis = Trim$(opis & " (" & lbl & ")")

                    moves.Add Array(d, brojDok, opis, ulaz, izlaz)'''
        new = '''                    Dim opis As String, opisKretanja As String
                    opis = NzToText(ambData(i, colTip))   ' TipAmbalaze
                    opisKretanja = AmbKarticaNazivKretanja(NzToText(ambData(i, 9)))
                    If Len(opisKretanja) = 0 Then opisKretanja = lbl
                    If Len(NzToText(ambData(i, 10))) > 0 Then
                        opisKretanja = "storno: " & opisKretanja
                    ElseIf CBool(ambData(i, 6)) Then
                        opisKretanja = opisKretanja & " - stornirano"
                    End If
                    If Len(opisKretanja) > 0 Then
                        opis = Trim$(opis & " (" & opisKretanja & ")")
                    End If

                    moves.Add Array(d, brojDok, opis, ulaz, izlaz)'''
        fn = replace_once(fn,old,new,'replace document label with actual movement kind')
        return fn
    s = in_function(s,'Public Function ReportKarticaAmbalaze(',change)
    return replace_once(s,'\nPrivate Function BuildOtkupBrojDokDict() As Object',
                        '\n' + LABEL_FUNCTION + '\nPrivate Function BuildOtkupBrojDokDict() As Object',
                        'insert movement label helper')


def decode_bas(raw: bytes) -> tuple[str,str]:
    try: return raw.decode('utf-8-sig').replace('\r\n','\n'),'utf-8'
    except UnicodeDecodeError: return raw.decode('cp1252').replace('\r\n','\n'),'cp1252'


def sha(raw: bytes) -> str: return hashlib.sha256(raw).hexdigest()


def git_show(repo: Path, path: str) -> bytes:
    c = subprocess.run(['git','-C',str(repo),'show',f'{BASE_REF}:src-vba/{path}'],
                       capture_output=True,check=False)
    if c.returncode:
        raise RuntimeError(f'Cannot read exact original {path}: {c.stderr.decode("utf8","replace").strip()}')
    return c.stdout


def run(repo: Path, mig18_dir: Path, out: Path, exported_amb: Path|None):
    if out.exists(): raise FileExistsError(f'Output exists, choose new --out-dir: {out}')
    old_report = (mig18_dir/'modIzvestaj.bas').read_bytes()
    old_amb = exported_amb.read_bytes() if exported_amb else git_show(repo,'modAmbalaza.bas')
    r, renc = decode_bas(old_report)
    a, aenc = decode_bas(old_amb)
    if 'Attribute VB_Name = "modIzvestaj"' not in r[:200]:raise ValueError('Not modIzvestaj.bas')
    if 'Attribute VB_Name = "modAmbalaza"' not in a[:200]:raise ValueError('Not modAmbalaza.bas')
    new_r = patch_izvestaj(r)
    new_a = patch_ambalaza(a)
    generated = {
      'modIzvestaj.bas':new_r.replace('\n','\r\n').encode(renc),
      'modAmbalaza.bas':new_a.replace('\n','\r\n').encode(aenc)
    }
    out.mkdir(parents=True)
    for name,raw in generated.items(): (out/name).write_bytes(raw)
    manifest = {
      'format':'AGRIx-MIG181-QA-AMBALAZA-MOVEMENT-LABELS-v1',
      'status':'GENERATED_BAS_ONLY_EXCEL_COMPILE_PENDING',
      'input_mig18_report_sha256':sha(old_report),
      'input_ambalaza_sha256':sha(old_amb),
      'source_ambalaza':str(exported_amb) if exported_amb else BASE_REF,
      'produced':{name:sha(raw) for name,raw in generated.items()},
      'workbook_modified':False,'repository_modified':False,
      'modifications':['append VrstaKretanja and StornoOd to AmbKretanjaNaloga (first 8 unchanged)',
                       'render movement label from VrstaKretanja; distinguish original and storno',
                       'retain MIG-18 chronological sorting + UI bypass unchanged'],
      'not_modified':['tblAmbalaza','financial data','workbook','modOtkup','modOtkupUI','transactional writers'],
      'excel_compile_pending':True,
    }
    (out/'MIG181_manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    print('MIG-18.1 LABEL MODULES GENERATED; Excel and repo untouched')
    for name,digest in manifest['produced'].items():print(name, digest)
    print('OUT:',out.resolve())


def main():
    p=argparse.ArgumentParser(description='Upgrade MIG-18 QA ambalaza descriptions, without changing ledger data')
    p.add_argument('--repo',default=str(Path.home()/'otkupapp-pwa'))
    p.add_argument('--mig18-modules',required=True,help='Folder holding MIG18-patched modIzvestaj.bas')
    p.add_argument('--out-dir',required=True)
    p.add_argument('--ambalaza-export',help='Optional freshly exported modAmbalaza.bas from QA workbook')
    x=p.parse_args()
    run(Path(x.repo),Path(x.mig18_modules),Path(x.out_dir),Path(x.ambalaza_export) if x.ambalaza_export else None)

if __name__=='__main__':main()
