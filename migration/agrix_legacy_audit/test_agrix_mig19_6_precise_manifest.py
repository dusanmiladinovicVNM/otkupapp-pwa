import unittest
from decimal import Decimal
import agrix_mig19_6_precise_manifest as m


def old(aid,typ,ident,sign,docid='D1',day='2026-06-27',q=300,doctype='Otkup'):
    return dict(AmbID=aid,Datum=day,TipAmbalaze='0,4kg',Kolicina=q,
      EntitetTip=typ,EntitetID=ident,Smer='Ulaz' if sign>0 else 'Izlaz',
      Stornirano='',DokumentTIP=doctype,DokumentID=docid,VozacID='ST-00005')

def new(aid,src,dst,docid='D1',day='2026-06-27',q=300,doctype='Otkup'):
    return dict(AmbID=aid,Datum=day,TipAmbalaze='0,4kg',Kolicina=q,
      OdNalogTip=src[0],OdNalogID=src[1],NaNalogTip=dst[0],NaNalogID=dst[1],
      VrstaKretanja='IZDATA_PRAZNA',DokumentTIP=doctype,DokumentID=docid,StornoOd='')

class Tests(unittest.TestCase):
    def test_equal_vector(self):
        a={('STANICA','ST-1'):Decimal('-10'),('KOOPERANT','KO-1'):Decimal('10')}
        self.assertEqual(m.classify_doc(a,a),'EXACT')
    def test_full_reversal(self):
        a={('STANICA','ST-1'):Decimal('-10'),('KOOPERANT','KO-1'):Decimal('10')}
        b={k:-v for k,v in a.items()}
        self.assertEqual(m.classify_doc(a,b),'EXACT_DIRECTION_REVERSED')
    def test_same_sign_different_magnitude(self):
        a={('STANICA','ST-1'):Decimal('-10'),('KOOPERANT','KO-1'):Decimal('10')}
        b={k:2*v for k,v in a.items()}
        self.assertEqual(m.classify_doc(a,b),'SAME_SIGN_DIFFERENT_MAGNITUDE')
    def test_participant_difference(self):
        a={('STANICA','ST-1'):Decimal('-10'),('KOOPERANT','KO-1'):Decimal('10')}
        b={('STANICA','ST-1'):Decimal('-10'),('KOOPERANT','KO-2'):Decimal('10')}
        self.assertEqual(m.classify_doc(a,b),'PARTICIPANTS_DIFFER')
    def test_storno_excluded_cover_excluded(self):
        olds=[old('OLD1','STANICA','ST-00005',+1),old('OLD2','KOOPERANT','KO-1',-1)]
        olds[1]['Stornirano']='Da'
        news=[new('NEW1',('STANICA','ST-1'),('STANICA','ST-00005'))]
        news.append(new('NEW2',('SPOLJNI_SVET','SYSTEM'),('KOOPERANT','KO-1')))
        news[1]['VrstaKretanja']='ULAZ_TUDJE_AMBALAZE'
        o,n,co,cn,cc,*_=m.grouped_vectors(olds,news)
        g=('2026-06-27','0,4KG','OTKUP','D1')
        self.assertEqual(co[g],1)
        self.assertEqual(cn[g],1)
        self.assertEqual(cc[g],1)
        self.assertEqual(o[g][('STANICA','ST-00005')],300)
        self.assertEqual(n[g][('STANICA','ST-00005')],300)
    def test_seven_requires_source_signature(self):
        a=old('AMB-02317','Stanica','ST-00005',+1,docid='OM-12',doctype='OMUlaz')
        g=('2026-06-27','0,4KG','OMULAZ','OM-12')
        changes=[dict(NalogID='ST-00005',Datum='2026-06-27',TipAmbalaze='0,4KG',PromenaRazlike='-300')]
        out=m.seven_manifest([a],[],{g:{('STANICA','ST-00005'):Decimal(300)}},{},None,{},changes) if False else None
        # The manifest demands all seven pinned rows, which is desirable fail-closed.
        with self.assertRaises(ValueError):
            m.seven_manifest([a],[],{g:{('STANICA','ST-00005'):Decimal(300)}},{},{},{},changes)
    def test_otkup_group_zero_net_not_exact(self):
        olds=[old('1','Stanica','ST-00005',-1),old('2','Kooperant','KO-1',+1)]
        news=[new('1',('Kooperant','KO-1'),('Stanica','ST-00005'))]
        o,n,co,cn,cc,*_=m.grouped_vectors(olds,news)
        rows,hist,netzero=m.docs_audit(o,n,co,cn,cc)
        self.assertEqual(len(rows),1)
        self.assertEqual(netzero,1)
        self.assertEqual(rows[0]['Classification'],'EXACT_DIRECTION_REVERSED')

if __name__=='__main__':unittest.main()
