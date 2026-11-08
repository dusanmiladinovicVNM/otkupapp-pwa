#!/usr/bin/env python3
"""No live private workbook required. Synthetic packaging+financial XLSM only."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import hashlib
import agrix_mig172_fix_novac_dates as m

M=m.MAIN;R=m.OFFICE;P=m.PKG

def xmlcol(x):
    z=''
    while x:z=chr(65+(x-1)%26)+z;x=(x-1)//26
    return z

def build(path,already=False):
    defs={
        'tblOtkup':(['OtkupID','Datum'],1702),
        'tblNovac':(['NovacID','Datum','Isplata'],39),
        'tblAmbalaza':(['AmbID','Datum'],4811),
        'tblAmbalazaDokument':(['AmbDokID','Datum'],48)
    }
    with ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{M}" xmlns:r="{R}"><sheets>'+''.join(f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>' for i in range(1,5))+'</sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{P}">'+''.join(f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml"/>' for i in range(1,5))+'</Relationships>')
        z.writestr('xl/styles.xml',f'<styleSheet xmlns="{M}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>')
        for i,(name,(cols,n)) in enumerate(defs.items(),1):
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'<Relationships xmlns="{P}"><Relationship Id="rId1" Target="../tables/table{i}.xml"/></Relationships>')
            tc=''.join(f'<tableColumn id="{j}" name="{x}"/>' for j,x in enumerate(cols,1))
            z.writestr(f'xl/tables/table{i}.xml',f'<table xmlns="{M}" name="{name}" displayName="{name}" ref="A1:{xmlcol(len(cols))}{n+1}"><tableColumns count="{len(cols)}">{tc}</tableColumns></table>')
            hd=''.join(f'<c r="{xmlcol(j)}1" t="inlineStr"><is><t>{c}</t></is></c>' for j,c in enumerate(cols,1))
            data=[]
            for rn in range(2,n+2):
                cells=[]
                for j,c in enumerate(cols,1):
                    ref=f'{xmlcol(j)}{rn}'
                    if c=='Datum':
                        st=' s="1"' if name=='tblOtkup' or (name=='tblNovac' and already) else ''
                        cells.append(f'<c r="{ref}"{st}><v>46214</v></c>')
                    elif c=='Isplata':cells.append(f'<c r="{ref}"><v>{rn}.5</v></c>')
                    else:cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{name}-{rn}</t></is></c>')
                data.append(f'<row r="{rn}">'+''.join(cells)+'</row>')
            z.writestr(f'xl/worksheets/sheet{i}.xml',f'<worksheet xmlns="{M}"><sheetData><row r="1">{hd}</row>'+''.join(data)+'</sheetData></worksheet>')
        z.writestr('xl/vbaProject.bin',b'PRESERVE VBA')

def run():
    with TemporaryDirectory() as td:
        tmp=Path(td);src=tmp/'synthetic.xlsm';build(src)
        out=tmp/'out'
        summ=m.run(src,out,m.sha(src))
        assert summ['finance_date_cells_corrected']==39
        assert summ['modified_parts']==['xl/worksheets/sheet2.xml']
        with ZipFile(src) as a,ZipFile(out/'AgriX_MIG17_2_NOVAC_DATES_QA.xlsm') as b:
            assert all(a.read(k)==b.read(k) for k in a.namelist() if k!='xl/worksheets/sheet2.xml')
            before=m.table_date_cells(a,m.find(a),'tblNovac')[1]
            after=m.table_date_cells(b,m.find(b),'tblNovac')[1]
            assert [t[2] for t in before]==[t[2] for t in after]
            assert len(before)==len(after)==39
            assert all(t[1]==1 for t in after)
        try:m.run(src,tmp/'out2','0'*64)
        except ValueError as e:assert 'SHA' in str(e)
        else:raise AssertionError('wrong SHA allowed')
        already=tmp/'already.xlsm';build(already,True)
        try:m.run(already,tmp/'out3',m.sha(already))
        except ValueError as e:assert 'already' in str(e)
        else:raise AssertionError('duplicate patch allowed')
    print('MIG-17.2 TESTS PASS: all 39 date styles, all other parts/VBA unchanged, wrong SHA blocked, already-patched blocked')
if __name__=='__main__':run()
