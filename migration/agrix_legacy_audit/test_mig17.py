from pathlib import Path
import tempfile,zipfile
from datetime import date
from decimal import Decimal
from xml.etree import ElementTree as ET
import agrix_mig17_import_ambalaza as m
from agrix_mig16_daily_close import compute_daily_policy

D=Decimal

def construct():
    dat=[('2026-08-01','Kooperant','K1','Stanica','S1','8','POVRAT_PRAZNE'),
         ('2026-08-02','Stanica','S1','Kooperant','K1','5','IZDATA_PRAZNA'),
         ('2026-08-03','Stanica','S1','Kooperant','K1','10','IZDATA_PRAZNA'),
         ('2026-08-04','Stanica','S1','Kooperant','K1','3','IZDATA_PRAZNA'),
         ('2026-08-04','Kooperant','K1','Stanica','S1','3','POVRAT_PRAZNE')]
    events=[]
    for i,(day,ft,fi,tt,ti,q,kind) in enumerate(dat):
        r={'Datum':day,'TipAmbalaze':'Gajbica','Kolicina':q,'DokumentID':f'OTK-{i}',
           'DokumentTIP':'OTKUP','OdNalogTip':ft,'OdNalogID':fi,
           'NaNalogTip':tt,'NaNalogID':ti,'VrstaKretanja':kind,'StornoOd':''}
        events.append({'case':str(i),'raw':{'target_column_values_WITHOUT_AMBID':r,'reserved_AmbID_NOT_ALLOCATED':f'AMB-X{i}'},'q':D(q),
                 'typ':'Gajbica','day':date.fromisoformat(day),'kind':kind,'from':(ft,fi),'to':(tt,ti)})
    return events

def test_daily():
    source=construct();daily,info=compute_daily_policy(source)
    r,p,counts,n=m.convert_physical(source,daily,'2026-10-21T12:00:00')
    assert n==5 and counts['foreign_cover_events']==1
    assert counts['relabel_return_events']==2
    assert len(r)==7
    assert sum(D(x['Kolicina']) for x in r if x['VrstaKretanja']=='ULAZ_TUDJE_AMBALAZE')==D(8)
    assert sum(D(x['Kolicina']) for x in r if x['VrstaKretanja']=='VRACANJE_TUDJE_AMBALAZE')==D(8)
    assert [x['historical_obligation_at_close'] for x in daily]==['8','3','0','0']
    try:m.convert_physical(source,[dict(daily[0], obligation_increase='9')]+daily[1:],'2026-10-21T12:00:00')
    except ValueError:pass
    else:raise AssertionError('Bad policy must fail')


def synthetic_workbook(path):
    M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    P='http://schemas.openxmlformats.org/package/2006/relationships'
    R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml',f'''<workbook xmlns="{M}" xmlns:r="{R}"><sheets><sheet name="Ambalaza" sheetId="1" r:id="rId1"/><sheet name="AmbalazaDokument" sheetId="2" r:id="rId2"/></sheets></workbook>''')
        z.writestr('xl/_rels/workbook.xml.rels',f'''<Relationships xmlns="{P}"><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>''')
        for i,(name,cols) in enumerate(m.COLS.items(),1):
            tablecols=''.join(f'<tableColumn id="{j}" name="{c}"/>' for j,c in enumerate(cols,1))
            last=m.col_name(len(cols))
            headers=''.join(f'<c r="{m.col_name(j)}1" t="inlineStr"><is><t>{c}</t></is></c>' for j,c in enumerate(cols,1))
            z.writestr(f'xl/worksheets/sheet{i}.xml',f'''<worksheet xmlns="{M}"><dimension ref="A1:{last}2"/><sheetData><row r="1">{headers}</row><row r="2"><c r="A2" s="1"/></row></sheetData></worksheet>''')
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'''<Relationships xmlns="{P}"><Relationship Id="rId1" Target="../tables/table{i}.xml"/></Relationships>''')
            z.writestr(f'xl/tables/table{i}.xml',f'''<table xmlns="{M}" name="{name}" displayName="{name}" ref="A1:{last}2"><autoFilter ref="A1:{last}2"/><tableColumns count="{len(cols)}">{tablecols}</tableColumns></table>''')
        z.writestr('xl/vbaProject.bin',b'VERIFIED-SYNTHETIC-VBA')

def test_write():
    events=construct();daily,_=compute_daily_policy(events)
    rows,_,_,_=m.convert_physical(events,daily,'2026-10-21T12:00:00')
    header=[{'AmbDokID':'ADK-0001','Vrsta':'REVERS','BrojDokumenta':'REV-1','Datum':'2026-08-01','BrojOwnerTip':'Stanica','BrojOwnerID':'S1','CreatedBy':'MIG-17','CreatedAt':'2026-10-21T12:00:00'}]
    with tempfile.TemporaryDirectory() as t:
        fn=Path(t)/'in.xlsm';synthetic_workbook(fn)
        with zipfile.ZipFile(fn) as z:
            patches=m.build_patches(z,{'tblAmbalaza':rows,'tblAmbalazaDokument':header})
            assert len(patches)==4
            assert any(b'IZDATA_PRAZNA' in v for v in patches.values())
            f=Path(t)/'out.xlsm'
            with zipfile.ZipFile(f,'w') as out:
                for i in z.infolist():out.writestr(i,patches.get(i.filename,z.read(i.filename)))
        with zipfile.ZipFile(f) as z:
            assert not z.testzip()
            m.check_rows(z,{'tblAmbalaza':len(rows),'tblAmbalazaDokument':len(header)})
            assert z.read('xl/vbaProject.bin')==b'VERIFIED-SYNTHETIC-VBA'
            for name in ['xl/worksheets/sheet1.xml','xl/tables/table1.xml','xl/worksheets/sheet2.xml','xl/tables/table2.xml']:
                ET.fromstring(z.read(name))
        print('MIG17 SYNTHETIC DAILY POLICY + TWO TABLE OOXML TESTS PASS')

if __name__=='__main__':
    test_daily()
    test_write()
