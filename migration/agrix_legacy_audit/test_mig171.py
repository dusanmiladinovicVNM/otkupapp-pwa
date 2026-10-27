from pathlib import Path
from decimal import Decimal
import tempfile,zipfile
import agrix_mig171_reconcile as m

M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
P='http://schemas.openxmlformats.org/package/2006/relationships'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def make(path):
    names={'tblAmbalaza':['AmbID','Datum','TipAmbalaze','Kolicina','OdNalogTip','OdNalogID','NaNalogTip','NaNalogID','VrstaKretanja','CreatedBy'],
           'tblNovac':['NovacID','KooperantID','OMID','Isplata','Datum','Stornirano'],
           'tblOtkup':['OtkupID','KooperantID','StanicaID','Stornirano'],
           'tblKooperanti':['KooperantID','StanicaID']}
    vals={
       'tblAmbalaza':{'AmbID':'A1','Datum':46235,'TipAmbalaze':'G','Kolicina':8,'OdNalogTip':'SpoljniSvet','NaNalogTip':'Kooperant','NaNalogID':'K1','VrstaKretanja':'ULAZ_TUDJE_AMBALAZE','CreatedBy':'MIG-17'},
       'tblNovac':{'NovacID':'N1','KooperantID':'K1','OMID':'ST-00001','Isplata':100,'Datum':46235,'Stornirano':''},
       'tblOtkup':{'OtkupID':'O1','KooperantID':'K1','StanicaID':'ST-00001','Stornirano':''},
       'tblKooperanti':{'KooperantID':'K1','StanicaID':'ST-00001'}}
    with zipfile.ZipFile(path,'w') as z:
        wb_sheets=''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i,name in enumerate(names,1))
        z.writestr('xl/workbook.xml',f'<workbook xmlns="{M}" xmlns:r="{R}"><sheets>{wb_sheets}</sheets></workbook>')
        wb_rels=''.join(f'<Relationship Id="rId{i}" Target="worksheets/sheet{i}.xml"/>' for i,_ in enumerate(names,1))
        z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{P}">{wb_rels}</Relationships>')
        for i,(name,cols) in enumerate(names.items(),1):
            header=''.join(f'<c r="{m.excelcol(j)}1" t="inlineStr"><is><t>{key}</t></is></c>' for j,key in enumerate(cols,1))
            body=''
            for j,key in enumerate(cols,1):
                v=vals[name].get(key,'')
                if v=='':continue
                coord=f'{m.excelcol(j)}2'
                if isinstance(v,(float,int)):
                    body+=f'<c r="{coord}"><v>{v}</v></c>'
                else:
                    body+=f'<c r="{coord}" t="inlineStr"><is><t>{v}</t></is></c>'
            z.writestr(f'xl/worksheets/sheet{i}.xml',f'<worksheet xmlns="{M}"><sheetData><row r="1">{header}</row><row r="2">{body}</row></sheetData></worksheet>')
            z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'<Relationships xmlns="{P}"><Relationship Id="rId1" Target="../tables/table{i}.xml"/></Relationships>')
            tablecols=''.join(f'<tableColumn id="{j}" name="{c}"/>' for j,c in enumerate(cols,1))
            z.writestr(f'xl/tables/table{i}.xml',f'<table xmlns="{M}" name="{name}" displayName="{name}" ref="A1:{m.excelcol(len(cols))}2"><tableColumns count="{len(cols)}">{tablecols}</tableColumns></table>')
        z.writestr('xl/vbaProject.bin',b'BINARY')


def run():
    m.excelcol=lambda n:chr(64+n)
    with tempfile.TemporaryDirectory() as t:
        p=Path(t)/'test.xlsm';make(p)
        book,formats=m.read_book(p,['tblAmbalaza','tblNovac','tblOtkup','tblKooperanti'])
        assert len(book['tblAmbalaza'])==1
        assert book['tblAmbalaza'][0]['Kolicina']=='8'
        assert formats['tblNovac']['not_date_format']==1
        assert m.day(book['tblNovac'][0]['Datum']) == m.EPOCH + m.timedelta(days=46235)
        detail,summary,totals=m.finance_diagnostics(book,'ST-00001')
        assert summary['in_report_scope_rows']==1
        assert summary['in_report_scope_positive_payment_rows']==1
        assert totals['in_report_scope_isplata']=='100'
        book['tblNovac'][0]['OMID']='ST-00002'
        assert m.finance_diagnostics(book,'ST-00001')[1]['in_report_scope_rows']==0
        assert m.norm('SpoljniSvet')=='spoljnisvet'
    from agrix_mig17_import_ambalaza import convert_physical
    from agrix_mig16_daily_close import compute_daily_policy
    from test_mig17 import construct
    cases=construct()
    daily,_=compute_daily_policy(cases)
    ledger,_,_,_=convert_physical(cases,daily,'2026-10-21T12:00:00')
    original_make=m.make_rows
    try:
        m.make_rows=lambda project: (cases,[],[],[],0,{})
        data={'tblAmbalaza':ledger,'tblOtkup':[{'StanicaID':'ST-00001','KooperantID':'K1'}]}
        recon,amb=m.amb_diagnostics(data,Path('.'),'ST-00001')
        assert amb['mismatch_pairs']==0 and amb['all_global_signed_parities']
        assert Decimal(recon[0]['new_physical'])-Decimal(recon[0]['mig17_cover'])==Decimal(recon[0]['old_signed_close'])
        ledger[0]['Kolicina']='9'
        recon2,amb2=m.amb_diagnostics(data,Path('.'),'ST-00001')
        assert amb2['mismatch_pairs']>0
    finally:m.make_rows=original_make
    print('MIG17.1 TESTS PASS: parser, date serial, payment filters, full AMB-cover reconciliation, imbalance detection')
if __name__=='__main__':run()
