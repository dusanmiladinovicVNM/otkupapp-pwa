import tempfile,json,re,sys
from pathlib import Path
from decimal import Decimal
from agrix_mig_archive import export_otkups,extract_issues,render_html,csv_save

def data():
 T={k:[] for k in ['tblOtkup','tblOtkupStavke','tblStanice','tblKooperanti','tblOtpremnica','tblOtpremnicaIzvori','tblOtpremnicaStavke','tblZbirna','tblZbirnaIzvori','tblZbirnaStavke','tblPrijemnica','tblPrijemnicaStavke','tblPaletaStavka']}
 T['tblStanice']=[{'StanicaID':'ST-00002','Naziv':'Stanica 2'}]
 T['tblKooperanti']=[{'KooperantID':'K1','Ime':'Petar','Prezime':'Petrovic'}]
 T['tblOtkup']=[{'OtkupID':'OTK-00001','Datum':'2026-06-17','BrojDokumenta':'A1','StanicaID':'ST-00002','KooperantID':'K1','VrstaVoca':'Malina','SortaVoca':'Willamette','Stornirano':''},
                   {'OtkupID':'OTK-00002','Datum':'2026-06-17','BrojDokumenta':'A2','StanicaID':'ST-00002','KooperantID':'K1','VrstaVoca':'Malina','Stornirano':'Da'},
                   {'OtkupID':'OTK-00003','Datum':'2026-06-17','BrojDokumenta':'A3','StanicaID':'ST-00002','KooperantID':'K1','VrstaVoca':'Malina','Stornirano':''}]
 T['tblOtkupStavke']=[{'OtkupStavkaID':'S1','OtkupID':'OTK-00001','Kolicina':'0.1','Cena':'120','Klasa':'I'},
                      {'OtkupStavkaID':'S2','OtkupID':'OTK-00001','Kolicina':'0.2','Cena':'120','Klasa':'II'},
                      {'OtkupStavkaID':'S3','OtkupID':'OTK-00002','Kolicina':'3','Cena':'80','Klasa':'I'}]
 T['tblOtpremnica']=[{'OtpremnicaID':'OTP-00666','Datum':'2026-07-05','IzdatoStatus':'IZDATO','Stornirano':'','StanicaID':'ST-00002','BrojOtpremnice':'A66'}]
 T['tblOtpremnicaStavke']=[{'OtpremnicaStavkaID':'PS1','OtpremnicaID':'OTP-00666','Kolicina':'10','KolAmbalaze':'1'}]
 T['tblZbirna']=[{'ZbirnaID':'ZBR-00666','Datum':'2026-07-05','Stornirano':'','BrojZbirne':'S2/050726'}]
 T['tblZbirnaStavke']=[{'ZbirnaStavkaID':'Z1','ZbirnaID':'ZBR-00666','Kolicina':'10'}]
 T['tblZbirnaIzvori']=[{'ZbirnaIzvorID':'ZI1','ZbirnaID':'ZBR-00666','OtpremnicaID':'OTP-00666'}]
 return T

def run():
 T=data();rows,extra=export_otkups(T)
 assert len(rows)==3
 assert rows[0]['kg']=='0.3',rows[0]
 assert rows[0]['value']=='36.0',rows[0]
 assert rows[1]['storno']
 assert rows[2]['kg'] is None
 issues=extract_issues(T,extra)
 assert any(x['code']=='ACTIVE_OTP_WITHOUT_ACTIVE_OTKUP' and x['document_id']=='OTP-00666' for x in issues),issues
 assert any(x['code']=='ARCHIVE_NEMA_STAVKI' for x in issues)
 h=render_html(rows,issues,{'source':'qa.xlsm','sha':'fake','generated':'test'})
 assert '<script id="archive-data"' in h
 assert 'OM · zbirno' in h and 'Kooperanti' in h and 'Izuzeci' in h
 d=json.loads(re.search(r'<script id="archive-data" type="application/json">(.*?)</script>',h,re.S).group(1))
 assert len(d['records'])==3 and len(d['issues'])>0
 with tempfile.TemporaryDirectory() as tmp:
  csv_save(Path(tmp)/'v.csv',[{'text':'=BAD()'},{'text':'+BAD'}],['text'])
  assert "'=BAD()" in (Path(tmp)/'v.csv').read_text('utf-8-sig')
 print('PASS: 8 tests (totals, storno, empty lines, OTP orphan, exception, HTML, JSON, CSV injection)')
if __name__=='__main__':run()
