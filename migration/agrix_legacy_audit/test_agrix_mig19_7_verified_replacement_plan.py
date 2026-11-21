import unittest
from decimal import Decimal
import agrix_mig19_7_verified_replacement_plan as m

def old(aid,entity,identifier,sign,qty=12,doc='OTK-01',storno=''):
    return dict(AmbID=aid,Datum='2026-06-23',TipAmbalaze='0,4kg',Kolicina=qty,
      DokumentTIP='Otkup',DokumentID=doc,EntitetTip=entity,EntitetID=identifier,
      Smer='Ulaz' if sign>0 else 'Izlaz',Stornirano=storno,VozacID='ST-2')

def new(aid,src=('Stanica','ST-2'),dst=('Kooperant','K-1'),qty=24,
        vk='AMBALAZA_UZ_ROBU',doc='OTK-01',storno=''):
    return dict(AmbID=aid,Datum='2026-06-23',TipAmbalaze='0,4kg',Kolicina=qty,
     OdNalogTip=src[0],OdNalogID=src[1],NaNalogTip=dst[0],NaNalogID=dst[1],
     DokumentTIP='Otkup',DokumentID=doc,VrstaKretanja=vk,StornoOd=storno)

def source():return [old('A1','Stanica','ST-2',-1),old('A2','Kooperant','K-1',1)]

class Tests(unittest.TestCase):
 def test_replace_wrong_quantity_exact_pair(self):
    rows,can,reason,types=m.get_otkup_proposals(source(),[new('N1')]);self.assertEqual(len(can),1)
    self.assertEqual(rows[0]['ExpectedQty'],'12')
    modified=m.rehearsal(source(),[new('N1')],can)
    acc,day,doc=m.reconcile(modified[0],modified[1],modified[2],modified[3],modified[4],modified[5]) if False else (None,None,None)
    self.assertEqual(modified[2][('KOOPERANT','K-1','0,4KG','2026-06-23')],Decimal(12))
    self.assertEqual(reason['ONE_TO_ONE_SOURCE_VECTOR_CANDIDATE'],1)
 def test_reverse_can_rehearse_when_kind_is_ambalaza(self):
    rows,can,_,_=m.get_otkup_proposals(source(),[new('N1',('Kooperant','K-1'),('Stanica','ST-2'),qty=12)])
    self.assertEqual(len(can),1)
    res=m.rehearsal(source(),[new('N1',('Kooperant','K-1'),('Stanica','ST-2'),qty=12)],can)
    self.assertEqual(res[2][('STANICA','ST-2','0,4KG','2026-06-23')],Decimal(-12))
 def test_invalid_direction_kind_blocked(self):
    s=[old('A1','Stanica','ST-2',1),old('A2','Kooperant','K-1',-1)]
    rows,can,reasons,_=m.get_otkup_proposals(s,[new('N1',qty=12,vk='IZDATA_PRAZNA')])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['EXISTING_MOVEMENT_KIND_INCOMPATIBLE_WITH_SOURCE_DIRECTION'],1)
 def test_unmatched_participant_blocked(self):
    rows,can,reasons,_=m.get_otkup_proposals(source(),[new('N1',dst=('Kooperant','K-2'))])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['TARGET_PARTICIPANT_IDENTITY_DIFFERENT'],1)
 def test_derived_cover_blocked(self):
    v=new('C',('SpoljniSvet','SVET'),('Kooperant','K-1'),qty=3,vk='ULAZ_TUDJE_AMBALAZE')
    rows,can,reasons,_=m.get_otkup_proposals(source(),[new('N1'),v])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['DERIVED_COVER_REQUIRES_INVARIANT_REVIEW'],1)
 def test_storno_link_blocked(self):
    rows,can,reasons,_=m.get_otkup_proposals(source(),[new('N1',storno='N0')])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['TARGET_HAS_STORNO_RELATION'],1)
 def test_referenced_original_storno_blocks(self):
    v=new('N2',('Stanica','ST-2'),('Kooperant','K-1'),doc='OTHER',storno='N1')
    rows,can,reasons,_=m.get_otkup_proposals(source(),[new('N1'),v])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['TARGET_HAS_STORNO_RELATION'],1)
 def test_multiple_target_rows_blocked(self):
    rows,can,reasons,_=m.get_otkup_proposals(source(),[new('N1',qty=8),new('N2',qty=16)])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['TARGET_NOT_SINGLE_PHYSICAL_ROW'],1)
 def test_unbalanced_original_blocked(self):
    src=[old('A1','Stanica','ST-2',-1),old('A2','Kooperant','K-1',1,qty=11)]
    rows,can,reasons,_=m.get_otkup_proposals(src,[new('N1')])
    self.assertEqual(len(can),0)
    self.assertEqual(reasons['SOURCE_LEGS_NOT_BALANCED'],1)
 def test_exact_document_skipped(self):
    rows,can,reason,kinds=m.get_otkup_proposals(source(),[new('N1',qty=12)])
    self.assertEqual(len(can),0)
    self.assertEqual(len(rows),0)
    self.assertEqual(kinds['EXACT'],1)
 def test_source_cancelled_removes_a_leg(self):
    src=source();src[1]['Stornirano']='Da'
    rows,can,reason,_=m.get_otkup_proposals(src,[new('N1')])
    self.assertEqual(len(can),0)
    self.assertEqual(reason['LEGACY_NOT_TWO_LEGS'],1)

if __name__=='__main__':unittest.main()
