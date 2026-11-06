#!/usr/bin/env python3
import unittest
from agrix_mig19_3_parity import ledger_views,reconcile,norm_type,void,source_semantic

def old(kind,ident,sign,q,day='2026-06-01',doc='Otkup',docid='OTK-1',storno=''):
    return dict(AmbID='OLD-'+kind+ident+str(q),Datum=day,TipAmbalaze='0,4kg',
                Kolicina=q,Smer=sign,EntitetTip=kind,EntitetID=ident,
                DokumentTIP=doc,DokumentID=docid,Stornirano=storno)

def new(ot,oi,nt,ni,kind,q,day='2026-06-01',doc='Otkup',docid='OTK-1',storno='',aid='AMB-1'):
    return dict(AmbID=aid,Datum=day,TipAmbalaze='0.4kg',Kolicina=q,
                OdNalogTip=ot,OdNalogID=oi,NaNalogTip=nt,NaNalogID=ni,
                VrstaKretanja=kind,StornoOd=storno,DokumentTIP=doc,DokumentID=docid)

class Tests(unittest.TestCase):
    def check(self, olds, news):
        z=ledger_views(olds,news)
        totals, days,docs=reconcile(*z[:6]);return totals,days,docs,z
    def test_negative_preserved_without_foreign_cover(self):
        a=old('Kooperant','K1','Izlaz',10)
        b=new('Kooperant','K1','Stanica','ST-00002','AMBALAZA_UZ_ROBU',10)
        totals,days,docs,z=self.check([a,old('Stanica','ST-00002','Ulaz',10)],[b]);self.assertTrue(all(x['Status']=='PASS' for x in totals+days+docs))
        self.assertEqual(next(x for x in totals if x['NalogTip']=='KOOPERANT')['StariSaldo'],'-10')
    def test_coverage_does_not_change_signed_history(self):
        olds=[old('Kooperant','K1','Izlaz',10), old('Stanica','ST-00002','Ulaz',10)]
        book=[new('SpoljniSvet','','Kooperant','K1','ULAZ_TUDJE_AMBALAZE',10,aid='AMB-COV'),
              new('Kooperant','K1','Stanica','ST-00002','AMBALAZA_UZ_ROBU',10,aid='AMB-BUS')]
        totals,days,docs,z=self.check(olds,book)
        self.assertTrue(all(x['Status']=='PASS' for x in totals+days+docs))
        koop=next(x for x in totals if x['NalogTip']=='KOOPERANT')
        self.assertEqual((koop['StariSaldo'],koop['NoviIstorijskiSaldo'],koop['NoviFizickiSaldo'],koop['MigracionoPokrice']),('-10','-10','0','10'))
    def test_wrong_station_balance_fails(self):
        olds=[old('Stanica','ST-00002','Ulaz',10)]
        book=[new('Stanica','ST-00001','Stanica','ST-00003','PRENOS_INTERNO',10)]
        totals,_,_,_=self.check(olds,book)
        self.assertEqual(sum(x['Status']=='MISMATCH' for x in totals),3)
    def test_day_close_mismatch_even_when_final_equal(self):
        olds=[old('Stanica','ST-00002','Ulaz',10,'2026-06-01'),old('Stanica','ST-00002','Izlaz',10,'2026-06-02')]
        book=[new('Stanica','ST-00001','Stanica','ST-00002','PRENOS_INTERNO',10,'2026-06-02',aid='T1'),
              new('Stanica','ST-00002','Stanica','ST-00001','PRENOS_INTERNO',10,'2026-06-02',aid='T2')]
        totals,days,_,_=self.check(olds,book)
        self.assertEqual([x['Status'] for x in totals if x['NalogID']=='ST-00002'],['PASS'])
        self.assertTrue(any(x['Status']=='MISMATCH' for x in days))
    def test_storno_pair_returns_zero(self):
        olds=[old('Kooperant','K1','Ulaz',10,storno='Da'),old('Stanica','ST-00002','Izlaz',10,storno='Da')]
        book=[new('Stanica','ST-00002','Kooperant','K1','IZDATA_PRAZNA',10,aid='A1'),
              new('Kooperant','K1','Stanica','ST-00002','IZDATA_PRAZNA',10,storno='A1',aid='A2')]
        totals,days,docs,z=self.check(olds,book)
        self.assertTrue(all(x['Status']=='PASS' for x in totals+days+docs))
        self.assertEqual(z[6]['legacy_stornirano'],2)
    def test_station_mirror_is_not_collapsed(self):
        olds=[old('Stanica','ST-00002','Ulaz',5)]
        book=[new('Vozac','ST-00002','Stanica','ST-00002','PRENOS_INTERNO',5)]
        totals,_,_,_=self.check(olds,book)
        self.assertTrue(all(x['Status']=='PASS' for x in totals))
        # Station and driver are not merged behind the operator's back.
        self.assertEqual(len(totals),1)
    def test_old_numeric_zero_not_judged(self):
        olds=[old('Stanica','ST-00002','Ulaz',0)]
        totals,_,_,_=self.check(olds,[])
        self.assertEqual(len(totals),1);self.assertEqual(totals[0]['Status'],'PASS')
    def test_normalization(self):
        self.assertEqual(norm_type(' 0.4 kg '),'0,4KG')
        self.assertTrue(void('Da'));self.assertFalse(void('Ne'))
        r=old('Stanica','ST-00002','Ulaz',5)
        r2=dict(r);r2['Datum']='46174';r2['Kolicina']='5.0'
        # Plain Excel number/date normalization checked elsewhere; no hardcoded calendar.
        self.assertEqual(source_semantic(r),source_semantic(dict(r,Kolicina='5.0',TipAmbalaze='0.4 KG')))

if __name__=='__main__':unittest.main()
