#!/usr/bin/env python3
"""MIG-ARCHIVE: prepare *source-matched*, narrow VBA patch; no workbook writes."""
import json
import difflib
import hashlib
from pathlib import Path
ROOT=Path(__file__).resolve().parent
PATCHES=[
    dict(module='modAmbalaza',function='AmbKretanjaNaloga',old='Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal iD As String) As Variant',new=[
        'Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal iD As String, _',
        '                                 Optional ByVal samoPoslovno As Boolean = False) As Variant',
    ]),
    dict(module='modAmbalaza',function='AmbKretanjaNaloga',old='    Dim cStorno As Long, cAmbID As Long',new=[
        '    Dim cStorno As Long, cAmbID As Long, cVrsta As Long',
    ]),
    dict(module='modAmbalaza',function='AmbKretanjaNaloga',old='    cAmbID = kolIdx(COL_AMB_ID)',new=[
        '    cAmbID = kolIdx(COL_AMB_ID)',
        '    cVrsta = kolIdx(COL_AMB_VRSTA_KRETANJA)',
    ]),
    dict(module='modAmbalaza',function='AmbKretanjaNaloga',old='            If znak <> 0 Then poz.Add Array(i, znak)',new=[
        '            If znak <> 0 Then',
        '                If samoPoslovno Then',
        '                    If Not AmbKarticaTehnickoPokrice(data, i, cVrsta, cStorno, vrsteDok, src) Then',
        '                        poz.Add Array(i, znak)',
        '                    End If',
        '                Else',
        '                    poz.Add Array(i, znak)',
        '                End If',
        '            End If',
    ]),
    dict(module='modAmbalaza',function='AmbKretanjaNaloga',old='End Function',new=[
        'End Function',
        '',
        "' MIG-ARCHIVE: only business-card projection skips technical deficit cover.",
        "' Storno of technical cover is also excluded; real transfers/storno remain.",
        'Private Function AmbKarticaTehnickoPokrice(ByRef data As Variant, ByVal i As Long, _',
        '                    ByVal cVrsta As Long, ByVal cStorno As Long, _',
        '                    ByVal vrsteDok As Object, ByVal sourceName As String) As Boolean',
        '    Dim vrsta As String, originalID As String',
        '    vrsta = AmbText(data(i, cVrsta))',
        '    originalID = AmbText(data(i, cStorno))',
        '    If Len(originalID) > 0 Then',
        '        If Not vrsteDok.Exists(originalID) Then',
        '            Err.Raise AMB_ERR_KNJIGA_KVAR, sourceName, _',
        '                      "StornoOd bez originala: " & originalID',
        '        End If',
        '        vrsta = CStr(vrsteDok(originalID))',
        '    End If',
        '    AmbKarticaTehnickoPokrice = (StrComp(vrsta, AMB_VK_ULAZ_TUDJE, vbTextCompare) = 0)',
        'End Function',
    ]),
    dict(module='modIzvestaj',function='ReportKarticaKooperanta',old='    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))',new=[
        '    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID), True)',
    ]),
    dict(module='modIzvestaj',function='ReportKarticaAmbalaze',old='    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))',new=[
        '    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID), True)',
    ]),
]

def locate(lines,fun,old):
    import re
    starts=[i for i,s in enumerate(lines) if re.match(r'^Public Function '+re.escape(fun)+r'\b',s)]
    if len(starts)!=1: raise ValueError(f'{fun} function starts: {starts}')
    start=starts[0]
    end=next((i for i in range(start+1,len(lines)) if lines[i].strip()=='End Function'),None)
    if end is None: raise ValueError(f'{fun} no End Function')
    occ=[i for i in range(start,end+1) if lines[i]==old]
    if len(occ)!=1: raise ValueError(f'{fun}: expected exactly one line {old!r}, got {len(occ)}')
    return occ[0]

def main():
    artifacts={}
    for name in ('modAmbalaza','modIzvestaj'):
        path=ROOT/'snapshot'/f'{name}.txt'
        old=path.read_text(encoding='utf-8-sig').splitlines()
        patched=old.copy()
        for p in sorted([x for x in PATCHES if x['module']==name],key=lambda x:locate(old,x['function'],x['old']),reverse=True):
            loc=locate(patched,p['function'],p['old'])
            patched[loc:loc+1]=p['new']
        (ROOT/'patched_preview'/f'{name}.bas').write_bytes((f'Attribute VB_Name = "{name}"\r\n'+'\r\n'.join(patched)+'\r\n').encode('utf-8'))
        delta=''.join(difflib.unified_diff([x+'\n' for x in old],[x+'\n' for x in patched],fromfile=f'{name}.ORIGINAL',tofile=f'{name}.PATCHED',lineterm='\n'))
        (ROOT/'patched_preview'/f'{name}.diff').write_text(delta,encoding='utf-8')
        artifacts[name]={'source_lines':len(old),'patched_lines':len(patched),'patches':sum(x['module']==name for x in PATCHES), 'sha256':hashlib.sha256(path.read_bytes()).hexdigest()}
    (ROOT/'patch_plan.json').write_text(json.dumps({'version':'MIG_ARCHIVE_KOOP_BUSINESS_V2','patches':PATCHES,'sources':artifacts},indent=2),encoding='utf-8')
    print(json.dumps(artifacts,indent=2))

if __name__=='__main__': main()
