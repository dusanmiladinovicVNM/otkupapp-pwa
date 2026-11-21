import unittest
from decimal import Decimal
from agrix_mig19_11_istorijska_semantika import (
    HISTORIC_DESCRIPTIONS, semantic_target, all_document_views,
    known_omulaz, firma_evidence, OMULAZ,
)


def old(dtyp,enttyp,entid,smer,amt,amb='AMB-00001',doc='OTK-00001',dt='2026-06-23'):
    return dict(AmbID=amb,Datum=dt,TipAmbalaze='0,4kg',Kolicina=str(amt),
        Smer=smer,EntitetTip=enttyp,EntitetID=entid,VozacID='',
        DokumentTIP=dtyp,DokumentID=doc,Stornirano='')


def target(kind,fromt,fromid,tot,toid,q,id,dtype='Otkup',doc='OTK-00001',dt='2026-06-23'):
    return dict(AmbID=id,Datum=dt,TipAmbalaze='0,4kg',Kolicina=str(q),
        OdNalogTip=fromt,OdNalogID=fromid,NaNalogTip=tot,NaNalogID=toid,
        DokumentTIP=dtype,DokumentID=doc,VrstaKretanja=kind,StornoOd='')


class HistoricalPolicyTests(unittest.TestCase):
    def setUp(self):
        self.source=[old('Otkup','Kooperant','K-1','Izlaz',10,'AMB-00001'),
             old('Otkup','Stanica','ST-00002','Ulaz',10,'AMB-00002'),
             old('OM-Izlaz-Koop','Kooperant','K-1','Ulaz',4,'AMB-00003'),
             old('OM-Izlaz-Koop','Stanica','ST-00002','Izlaz',4,'AMB-00004')]
        self.new=[target('AMBALAZA_UZ_ROBU','Kooperant','K-1','Stanica','ST-00002',10,'M-0001'),
            target('IZDATA_PRAZNA','Stanica','ST-00002','Kooperant','K-1',4,'M-0002')]

    def test_nine_exact_historical_descriptions(self):
        self.assertEqual(len(HISTORIC_DESCRIPTIONS),9)
        self.assertIn('OMULAZ',HISTORIC_DESCRIPTIONS)

    def test_read_model_relabels_only_issuance(self):
        projection,relabel,ungrounded=semantic_target(self.source,self.new)
        self.assertEqual(relabel,['M-0002'])
        self.assertFalse(ungrounded)
        self.assertEqual(projection[0]['DokumentTIP'],'Otkup')
        self.assertEqual(projection[1]['DokumentTIP'],'OM-Izlaz-Koop')
        self.assertEqual(self.new[1]['DokumentTIP'],'Otkup')

    def test_full_document_parity_with_two_typed_kinds(self):
        projection,_,_=semantic_target(self.source,self.new)
        differences,by_kind,_,_,_=all_document_views(self.source,projection)
        self.assertEqual(differences,[])
        self.assertEqual(by_kind,{})
        original_diff,orig_kind,_,_,_=all_document_views(self.source,self.new)
        self.assertGreater(len(original_diff),0)

    def test_no_invented_label_without_old_signature(self):
        projected,relab,held=semantic_target(self.source[:2],self.new)
        self.assertEqual(len(relab),0)
        self.assertEqual(held,['M-0002'])
        self.assertEqual(projected[1]['DokumentTIP'],'Otkup')

    def test_does_not_relabel_wrong_quantity(self):
        wrong=target('IZDATA_PRAZNA','Stanica','ST-00002','Kooperant','K-1',5,'M-0002')
        projected,relab,held=semantic_target(self.source,[self.new[0],wrong])
        self.assertEqual(relab,[])
        self.assertEqual(held,['M-0002'])

    def test_does_not_relabel_wrong_partner(self):
        wrong=target('IZDATA_PRAZNA','Stanica','ST-00002','Kooperant','K-2',4,'M-0002')
        projected,relab,held=semantic_target(self.source,[self.new[0],wrong])
        self.assertEqual(relab,[])
        self.assertEqual(held,['M-0002'])

    def test_cover_stays_separate(self):
        cover=target('ULAZ_TUDJE_AMBALAZE','SpoljniSvet','SpoljniSvet','Kooperant','K-1',7,'M-0003')
        projected,relab,_=semantic_target(self.source,self.new+[cover])
        self.assertEqual(relab,['M-0002'])
        self.assertEqual(projected[-1]['VrstaKretanja'],'ULAZ_TUDJE_AMBALAZE')
        mismatch,_,_,_,_=all_document_views(self.source,projected)
        self.assertEqual(mismatch,[])

    def test_all_six_as_given(self):
        src=[]
        for amb,(d,st,q,origin) in OMULAZ.items():
            src.append(old('OMUlaz','Stanica',st,'Ulaz',q,amb,amb,dt=d))
        rows,issues=known_omulaz(src)
        self.assertFalse(issues)
        self.assertEqual(sum(x['Prevod']=='PRENOS_INTERNO' for x in rows),3)
        self.assertEqual(sum(x['Status']=='HOLD_IZVOR_NEPOZNAT' for x in rows),3)
        self.assertEqual(sum(Decimal(x['Kolicina']) for x in rows if x['Prevod']=='PRENOS_INTERNO'),Decimal(633))
        self.assertEqual(sum(Decimal(x['Kolicina']) for x in rows if x['Status']=='HOLD_IZVOR_NEPOZNAT'),Decimal(18))

    def test_omulaz_strict_source_signature(self):
        src=[old('OMUlaz','Stanica','ST-00005','Izlaz',300,'AMB-02317',dt='2026-06-27')]
        rows,issues=known_omulaz(src)
        self.assertTrue(any('direction' in e for e in issues))

    def test_firma_does_not_fill_from_mirror(self):
        r=old('OM-Ulaz-Firma','Stanica','ST-00005','Ulaz',4)
        r['VozacID']='ST-00005'
        f=firma_evidence([r])[0]
        self.assertEqual(f['VozacID_ZAPIS'],'ST-00005')
        self.assertEqual(f['Izvor_IzIstorijskogZapisa'],'')
        self.assertEqual(f['Odrediste_IzIstorijskogZapisa'],'')

    def test_reverse_split_can_sum_source(self):
        # Canonical two reverse legs are one legacy issuance event, split by ownership.
        r1=target('IZDATA_PRAZNA','Stanica','ST-00002','Kooperant','K-1',2,'M-0002')
        r2=target('VRACANJE_TUDJE_AMBALAZE','Stanica','ST-00002','Kooperant','K-1',2,'M-0003')
        projected,labels,held=semantic_target(self.source,[self.new[0],r1,r2])
        self.assertEqual(len(labels),2)
        diffs,_,_,_,_=all_document_views(self.source,projected)
        self.assertEqual(diffs,[])


if __name__=='__main__': unittest.main()
