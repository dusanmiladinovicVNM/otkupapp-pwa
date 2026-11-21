import tempfile
from pathlib import Path
import json
import unittest
from legacy_audit import Table, AuditError
from legacy_stage_v2 import make_plan, transform_one, ensure_inputs, deterministic_id, write_private_jsonl


def t(name,rows):
    columns=list(rows[0].keys()) if rows else []
    return Table(name,name,columns,rows)

SCHEMA={
 'tblOtkup':['OtkupID','Datum','KooperantID','StanicaID','OtpremnicaID','BrojZbirne'],
 'tblOtkupStavke':['OtkupStavkaID','OtkupID','RedniBroj','Klasa','Kolicina','Cena','KolAmbalaze','BrutoKg'],
 'tblOtpremnica':['OtpremnicaID','StanicaID','BrojZbirne'],
 'tblOtpremnicaStavke':['OtpremnicaStavkaID','OtpremnicaID','RedniBroj','Klasa','Kolicina','KolAmbalaze','BrutoKg','PredlogCena'],
 'tblZbirna':['ZbirnaID','BrojZbirne','UkupnoKolicina','UkupnoAmbalaze','Klasa','VrstaVoca'],
 'tblZbirnaStavke':['ZbirnaStavkaID','ZbirnaID','RedniBroj','Klasa','Kolicina','KolAmbalaze'],
}
SCHEMA.update({k:[v] for k,v in [('tblMagacin','MagacinID'),('tblPrijemnica','PrijemnicaID'),
('tblPaleta','PaletaID'),('tblPaletaStavka','StavkaID'),('tblPrerada','PreradaID'),('tblPreradaStavka','StavkaID')]})
for k in ['tblArtikli','tblKooperanti','tblStanice','tblKupci','tblVozaci','tblKulture','tblParcele',
  'tblTipAmbalaze','tblTipPalete','tblKutije','tblKese','tblVrstaGotovihProizvoda','tblCenovnik']:
    SCHEMA[k]=[]

def fixture():
    x=[t('tblOtkup',[{'OtkupID':'OTK-1','Datum':'45678','KooperantID':'KOOP-1','StanicaID':'STA-1',
                      'Kolicina':'101.5','Cena':'40','Klasa':'I','KolAmbalaze':'10','BrutoKg':'110',
                      'OtpremnicaID':'OTP-1','BrojZbirne':'ZBR-BROJ-1'}]),
       t('tblOtpremnica',[{'OtpremnicaID':'OTP-1','StanicaID':'STA-1','BrojZbirne':'ZBR-BROJ-1',
                           'Kolicina':'101.5','Cena':'50','Klasa':'I','KolAmbalaze':'10','BrutoKg':'110'}]),
       t('tblZbirna',[{'ZbirnaID':'ZBR-1','BrojZbirne':'ZBR-BROJ-1','UkupnoKolicina':'101.5',
                      'UkupnoAmbalaze':'10','Klasa':'I','VrstaVoca':'fruit'}]),
       t('tblPrijemnica',[{'PrijemnicaID':'PRJ-1','BrojZbirne':'ZBR-BROJ-1'}]),
       t('tblPaleta',[{'PaletaID':'PAL-1'}]),
       t('tblPaletaStavka',[{'StavkaID':'PS-1','PaletaID':'PAL-1','PrijemnicaID':'PRJ-1'}]),
       t('tblPrerada',[{'PreradaID':'PRD-1'}]),
       t('tblPreradaStavka',[{'StavkaID':'PRS-1','PreradaID':'PRD-1','PaletaID':'PAL-1'}]),
       t('tblMagacin',[{'MagacinID':'MAG-1','ArtikalID':'ART-POC-DUG','Tip':'Izlaz'}]),
       t('tblAmbalaza',[{'AmbID':'AMB-1','DokumentID':'OTK-1','DokumentTip':'Otkup'}])]
    return {a.name.lower():a for a in x}

class StageTests(unittest.TestCase):
 def test_one_to_one(self):
    s,r=make_plan(fixture(),SCHEMA,'aa')
    for name in ('Otkup','Otpremnica','Zbirna'):
       self.assertEqual(len(s['tbl'+name]),1)
       self.assertEqual(len(s['tbl'+name+'Stavke']),1)
    self.assertEqual(s['tblOtkupStavke'][0]['Kolicina'],'101.5')
    self.assertEqual(s['tblOtpremnicaStavke'][0]['PredlogCena'],'50')
    self.assertEqual(s['tblZbirnaStavke'][0]['Kolicina'],'101.5')
    self.assertNotIn('UkupnoKolicina',s['tblZbirna'][0])
    self.assertNotIn('BrojZbirne',s['tblOtkup'][0])
    self.assertEqual(r['exact_id_links']['paleta_item_to_receipt']['present'],1)
    self.assertEqual(r['exact_id_links']['prerada_item_to_paleta']['present'],1)
 def test_no_ambalaza_import(self):
    s,r=make_plan(fixture(),SCHEMA,'aa')
    self.assertNotIn('tblAmbalaza',s)
    self.assertEqual(len(s['UNCONVERTED__tblAmbalaza']),1)
 def test_missing_parent_reported(self):
    d=fixture(); d['tblprijemnica'].rows[0]['BrojZbirne']='NOT-FOUND'
    _,report=make_plan(d,SCHEMA,'aa')
    self.assertEqual(report['missing_zbirna_by_business_number_ordinals']['tblPrijemnica'],[1])
 def test_refusal_of_duplicate_id(self):
    d=fixture();d['tblotkup'].rows.append(dict(d['tblotkup'].rows[0]))
    with self.assertRaises(AuditError):make_plan(d,SCHEMA,'aa')
 def test_refusal_nonopening_magacin(self):
    d=fixture();d['tblmagacin'].rows[0]['ArtikalID']='AR-REAL'
    with self.assertRaises(AuditError):make_plan(d,SCHEMA,'aa')
 def test_deterministic_item_id(self):
    self.assertEqual(deterministic_id('a','tblOtkupStavke','OTK-1'),deterministic_id('a','tblOtkupStavke','OTK-1'))
    self.assertNotEqual(deterministic_id('a','tblOtkupStavke','OTK-1'),deterministic_id('b','tblOtkupStavke','OTK-1'))
 def test_no_private_values_in_manifest(self):
    s,r=make_plan(fixture(),SCHEMA,'aa')
    summary=json.dumps(r)
    self.assertNotIn('OTK-1',summary)
    self.assertNotIn('KOOP-1',summary)
    self.assertNotIn('fruit',summary)
 def test_no_overwrite_stage(self):
    with tempfile.TemporaryDirectory() as td:
       p=Path(td)/'f.jsonl'
       write_private_jsonl(p,[{'X':'first'}])
       with self.assertRaises(FileExistsError):write_private_jsonl(p,[{'X':'second'}])
       self.assertIn('first',p.read_text())

if __name__=='__main__':unittest.main()
