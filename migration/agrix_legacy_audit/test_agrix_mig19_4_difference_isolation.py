import unittest
from collections import defaultdict
from decimal import Decimal

from agrix_mig19_4_difference_isolation import (gate_sum, day_isolation, six_checks,
                                                   document_statistics, mismatch_category_station_gaps)


def old(aid,station,date='2026-06-27',qty=300,pack='0,4kg'):
    return {'AmbID':aid,'Datum':date,'TipAmbalaze':pack,'Kolicina':qty,
            'EntitetTip':'Stanica','EntitetID':station,'Smer':'Ulaz',
            'DokumentTIP':'OMUlaz','DokumentID':'','Stornirano':''}

def canonical(station,source='ST-00001',date='2026-06-27',qty=300,pack='0,4kg'):
    return {'AmbID':'NEW-1','Datum':date,'TipAmbalaze':pack,'Kolicina':qty,
            'OdNalogTip':'Stanica','OdNalogID':source,'NaNalogTip':'Stanica',
            'NaNalogID':station,'VrstaKretanja':'PRENOS_INTERNO','StornoOd':'',
            'DokumentTIP':'AmbalazaDokument','DokumentID':'ADK-1'}

class MIG194Tests(unittest.TestCase):
    def test_conservation_nonzero_not_fixed_by_internal_transfers(self):
        pairs=[dict(NalogTip='STANICA',NalogID='ST-00002',TipAmbalaze='0,4KG',
                    StariSaldo='83',NoviIstorijskiSaldo='0',Razlika='-83',Status='MISMATCH')]
        gaps,conserv=gate_sum(pairs)
        self.assertEqual(conserv[0]['ZbirNoviMinusStari'],'-83')
        self.assertEqual(conserv[0]['MozeLiSamoInterniPrenosIzravnatiSvaStanja'],'NE')
    def test_conservation_zero_is_not_a_proof(self):
        pairs=[]
        for station,diff in [('ST-00001','-83'),('ST-00002','83')]:
            pairs.append(dict(NalogTip='STANICA',NalogID=station,TipAmbalaze='0,4KG',
                    StariSaldo='0',NoviIstorijskiSaldo=diff,Razlika=diff,Status='MISMATCH'))
        _,conserv=gate_sum(pairs)
        self.assertEqual(conserv[0]['ZbirNoviMinusStari'],'0')
        self.assertEqual(conserv[0]['MozeLiSamoInterniPrenosIzravnatiSvaStanja'],'NEODREDJENO')
    def test_daily_first_divergence(self):
        old=defaultdict(Decimal,{('STANICA','ST-00002','0,4KG','2026-06-01'):Decimal(5)})
        target=defaultdict(Decimal,{('STANICA','ST-00002','0,4KG','2026-06-02'):Decimal(5)})
        tracked=[dict(NalogTip='STANICA',NalogID='ST-00002',TipAmbalaze='0,4KG')]
        increments,first,off=day_isolation(old,target,tracked)
        self.assertEqual(len(increments),2)
        self.assertEqual(first[0]['PrvaRazlikaDatum'],'2026-06-01')
        self.assertEqual(first[0]['ZavrsnaRazlika'],'0')
        self.assertEqual(off[('STANICA','ST-00002','0,4KG')],1)
    def test_six_source_candidate_is_not_proof(self):
        src=[old('AMB-02317','ST-00005')]
        current=[canonical('ST-00005')]
        known,res=six_checks(src,current,[])
        hit=[r for r in known if r['LegacyAmbID']=='AMB-02317'][0]
        self.assertEqual(hit['ST00001InternalTransferCandidates'],1)
        self.assertTrue(hit['SourceFieldsMatchExpected'])
        self.assertFalse(hit['LineageProven'])
        self.assertEqual(res[0]['LegacySixInputSubtotal'],'300')
    def test_documents_old_only_new_only_and_both_wrong(self):
        a=('STANICA','ST-00002','0,4KG','2026-06-01','OTKUP','OTK-1')
        b=('STANICA','ST-00002','0,4KG','2026-06-01','OTKUP','OTK-2')
        c=('STANICA','ST-00002','0,4KG','2026-06-01','OTKUP','OTK-3')
        rows=document_statistics({a:Decimal(5),c:Decimal(5)},{b:Decimal(7),c:Decimal(4)})
        r=rows[0]
        self.assertEqual((r['Razlicitih'],r['StariSamo'],r['NoviSamo'],r['ObaRazlicita']),(3,1,1,1))
    def test_breakdown_preserves_negative_numbers(self):
        rows=mismatch_category_station_gaps([old('AMB-02317','ST-00005')],
            [canonical('ST-00005')],[dict(NalogID='ST-00005',TipAmbalaze='0,4KG')],'2026-08-31')
        self.assertTrue(any(r['Model']=='OLD' and r['PotpisaniEfekat']=='300' for r in rows))
        self.assertTrue(any(r['Model']=='NEW' and r['PotpisaniEfekat']=='300' for r in rows))

if __name__=='__main__':unittest.main()
