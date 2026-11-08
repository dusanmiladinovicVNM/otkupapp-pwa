import unittest
from decimal import Decimal
from agrix_mig19_9_full_document_shadow_rehearsal import (
    doc_reconstruction_plan, full_source_projection_shadow, count_parity
)

D='2026-06-22';T='0,4KG';DOC='OTK-0001'
S=('STANICA','ST-00003');K=('KOOPERANT','K-00001')

def old(a,sgn,q='10',aid='',doc=DOC):
    return dict(AmbID=aid,Datum=D,TipAmbalaze=T,DokumentTIP='Otkup',DokumentID=doc,
        EntitetTip=a[0],EntitetID=a[1],Kolicina=q,Smer='Ulaz' if sgn>0 else 'Izlaz',Stornirano='Ne')

def new(a,b,q,kind='AMBALAZA_UZ_ROBU',aid='N1',doc=DOC,storno=''):
    return dict(AmbID=aid,Datum=D,TipAmbalaze=T,DokumentTIP='Otkup',DokumentID=doc,
        OdNalogTip=a[0],OdNalogID=a[1],NaNalogTip=b[0],NaNalogID=b[1],
        Kolicina=str(q),VrstaKretanja=kind,StornoOd=storno)

src=[old(S,-1,aid='A1'),old(K,1,aid='A2')]

class Tests(unittest.TestCase):
    def test_exact_untouched(self):
        g,a,k,s=doc_reconstruction_plan(src,[new(S,K,10)])
        self.assertEqual(len(k),0);self.assertEqual(len(g),0)

    def test_forward_reverse_mismatch_document(self):
        dst=[new(S,K,10,aid='N1',kind='IZDATA_PRAZNA'),new(K,S,2,aid='N2')]
        g,a,k,s=doc_reconstruction_plan(src,dst)
        self.assertEqual(len(g),1)
        self.assertEqual(g[0]['Reason'],'SOURCE_EXACT_FORWARD_PLUS_REVERSE')
        self.assertEqual(g[0]['ForwardQuantity'],'10')
        self.assertEqual(g[0]['ReverseQuantity'],'2')
        self.assertEqual(s['exact_same_direction_anchor_count'],1)
        self.assertEqual(len(a),2)

    def test_shadow_matches_source_and_is_invalid_kind(self):
        dst=[new(S,K,10,aid='N1'),new(K,S,2,aid='N2')]
        g,a,keys,_=doc_reconstruction_plan(src,dst)
        after,removed,added=full_source_projection_shadow(src,dst,keys)
        self.assertEqual((removed,added),(2,1))
        self.assertEqual(after[0]['VrstaKretanja'],'__INVALID_SHADOW_KIND_NOT_AMB10__')
        self.assertEqual(count_parity(src,after)['otkup_document_account_slice_mismatches'],0)
        self.assertEqual(count_parity(src,dst)['otkup_document_account_slice_mismatches'],2)

    def test_cover_preserved_but_flagged(self):
        dst=[new(S,K,10,aid='N1'),new(K,S,2,aid='N2'),
             new(('SPOLJNISVET','OUT'),K,4,kind='ULAZ_TUDJE_AMBALAZE',aid='C1')]
        g,a,keys,s=doc_reconstruction_plan(src,dst)
        self.assertEqual(s['docs_with_derived_cover'],1)
        after,removed,added=full_source_projection_shadow(src,dst,keys)
        self.assertEqual(len(after),2)
        self.assertEqual({r['AmbID'] for r in after if r['VrstaKretanja']=='ULAZ_TUDJE_AMBALAZE'}, {'C1'})
        self.assertEqual(count_parity(src,after)['otkup_document_account_slice_mismatches'],0)

    def test_storno_link_flagged(self):
        dst=[new(S,K,10,aid='N1'),new(K,S,2,aid='N2',storno='N1')]
        g,a,keys,s=doc_reconstruction_plan(src,dst)
        self.assertEqual(s['docs_with_storno_dependency'],1)
        self.assertEqual(len(keys),1) # shadow can still prove math; NOT an approved writer

    def test_previously_exact_plus_bad_do_not_conflate(self):
        oldrows=src+[old(S,-1,aid='A3',doc='OTK-0002'),old(K,1,aid='A4',doc='OTK-0002')]
        newrows=[new(S,K,10,aid='N1'),new(K,S,2,aid='N2'),new(S,K,10,aid='N3',doc='OTK-0002')]
        g,a,keys,s=doc_reconstruction_plan(oldrows,newrows)
        self.assertEqual(len(keys),1)
        after,removed,added=full_source_projection_shadow(oldrows,newrows,keys)
        self.assertEqual(len(after),2)
        self.assertEqual(count_parity(oldrows,after)['otkup_document_account_slice_mismatches'],0)

    def test_incompatible_target_endpoint_is_blocked(self):
        dst=[new(S,K,10,aid='N1'),new(S,('KOOPERANT','K-OTHER'),2,aid='N2')]
        with self.assertRaisesRegex(ValueError,'Unexpected group shape'):
            doc_reconstruction_plan(src,dst)

    def test_no_source_side_for_orphan_blocked(self):
        with self.assertRaisesRegex(ValueError,'Legacy pair not proven'):
            doc_reconstruction_plan([], [new(S,K,10),new(K,S,3,aid='N2')])

if __name__=='__main__':unittest.main()
