import unittest
from decimal import Decimal
from agrix_mig19_8_multi_event_otkup import evaluate, analysis

G=('2026-06-22','0,4KG','OTKUP','OTK-001')

def leg(a,way,qty='10',aid='AMB-00001'):
    return {'EntitetTip':a[0],'EntitetID':a[1], 'Smer':way,'Kolicina':qty,'AmbID':aid,'Datum':G[0],
       'TipAmbalaze':G[1],'DokumentTIP':G[2],'DokumentID':G[3],'Stornirano':'Ne'}
def tx(a,b,qty='10',kind='IZDATA_PRAZNA'):
    return {'OdNalogTip':a[0],'OdNalogID':a[1],'NaNalogTip':b[0],'NaNalogID':b[1],
        'Kolicina':qty,'VrstaKretanja':kind,'Datum':G[0], 'TipAmbalaze':G[1],
        'DokumentTIP':G[2],'DokumentID':G[3]}
A=('STANICA','ST-00003');B=('KOOPERANT','K-100')
OLD=[(1,leg(A,'Izlaz','10','AMB-00001')),(2,leg(B,'Ulaz','10','AMB-00002'))]
class Check(unittest.TestCase):
 def test_exact_single(self):
    res=evaluate(G,OLD,[(1,tx(A,B))],[])
    self.assertEqual(res['Status'],'FULL_VECTOR_EXACT')
    self.assertEqual(res['TargetNonCoverRows'],1)
 def test_exact_split(self):
    res=evaluate(G,OLD,[(1,tx(A,B,'6')),(2,tx(A,B,'4'))],[])
    self.assertEqual(res['Status'],'FULL_VECTOR_EXACT')
    self.assertTrue(res['SameDirectionSplitEqualsOriginal'])
 def test_reverse(self):
    res=evaluate(G,OLD,[(1,tx(B,A))],[])
    self.assertEqual(res['Status'],'SAME_PAIR_WITH_REVERSE_COMPONENTS')
    self.assertEqual(res['FromDelta'],'20')
 def test_same_pair_partial(self):
    res=evaluate(G,OLD,[(1,tx(A,B,'7')),(2,tx(A,B,'2'))],[])
    self.assertEqual(res['ToDelta'],'-1')
    self.assertEqual(res['Status'],'SAME_PAIR_QUANTITY_OR_DIRECTION_MISMATCH')
 def test_extra_actor(self):
    res=evaluate(G,OLD,[(1,tx(A,('VOZAC','V-1'),'10')),(2,tx(('VOZAC','V-1'),B,'10'))],[])
    self.assertEqual(res['Status'],'FULL_VECTOR_EXACT')
    self.assertEqual(res['CanonicalTouchesOtherAccounts'],1)
 def test_broken_bridge(self):
    res=evaluate(G,OLD,[(1,tx(A,('VOZAC','V-1'),'10')),(2,tx(('VOZAC','V-1'),B,'8'))],[])
    self.assertEqual(res['Status'],'EXTRA_CANONICAL_PARTICIPANTS')
 def test_cover_excluded(self):
    res=evaluate(G,OLD,[(1,tx(A,B))],[(2,tx(('SPOLJNI_SVET','X'),A,'10','ULAZ_TUDJE_AMBALAZE'))])
    self.assertEqual(res['TargetCoverRows'],1)
    self.assertTrue(res['VectorCompletelyEqual'])
 def test_unpaired_source(self):
    res=evaluate(G,OLD[:1],[(1,tx(A,B))],[])
    self.assertFalse(res['SourcePairProven'])
    self.assertIn('SOURCE_PAIR_NOT_PROVEN',res['Status'])
 def test_analysis(self):
    old=[x[1] for x in OLD]; new=[tx(A,B,'6'),tx(A,B,'4')]
    rows,ledger,s=analysis(old,new)
    self.assertEqual(s['whole_doc_vector_exact'],1)
    self.assertEqual(s['target_noncover_row_count_distribution'],{2:1})
    self.assertEqual(len(ledger),2)
if __name__=='__main__':unittest.main()
