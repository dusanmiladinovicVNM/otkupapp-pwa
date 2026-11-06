"""Synthetic MIG-01.3 tests. All values are fake; no customer data."""
import unittest
from legacy_audit import Table, AuditError
from legacy_production_scope import evaluate, parse_scope, KEPT, TEST, REVIEW, UNKNOWN


def table(name, col, rows):
    return Table(name,name,col,rows)

class ProductionScopeTests(unittest.TestCase):
    def setUp(self):
        self.tabs={
            'tblotkup':table('tblOtkup',['OtkupID','KooperantID','StanicaID','BrojZbirne','BrojDokumenta','Datum','KulturaID'],[
                {'OtkupID':'OTK-1','KooperantID':'KOP-1','StanicaID':'STA-1','BrojZbirne':'ZBR-MISSING','BrojDokumenta':'1','Datum':'1','KulturaID':'KUL-1'},
                {'OtkupID':'OTK-TEST','KooperantID':'KOP-TEST','StanicaID':'STA-1','BrojZbirne':'ZBR-1','BrojDokumenta':'2','Datum':'1','KulturaID':'KUL-1'}]),
            'tblkooperanti':table('tblKooperanti',['KooperantID'],[
                {'KooperantID':'KOP-1'},{'KooperantID':'KOP-TEST'}]),
            'tblzbirna':table('tblZbirna',['ZbirnaID','GeneracijaID','BrojZbirne','VozacID','KupacID','Datum','Klasa'],[
                {'ZbirnaID':'ZBR-ID1','GeneracijaID':'GEN-1','BrojZbirne':'ZBR-1','VozacID':'VOZ','KupacID':'KUP','Datum':'D','Klasa':'I'},
                {'ZbirnaID':'ZBR-ID2','GeneracijaID':'GEN-1','BrojZbirne':'ZBR-1','VozacID':'VOZ','KupacID':'KUP','Datum':'D','Klasa':'II'}]),
            'tblnovac':table('tblNovac',['NovacID','OtkupID'],[
                {'NovacID':'NOV-1','OtkupID':'OTK-1'},
                {'NovacID':'NOV-2','OtkupID':'OTK-TEST'},
                {'NovacID':'NOV-3','OtkupID':'OTK-TEST'}]),
            'tblutovar':table('tblUtovar',['UtovarID'],[{'UtovarID':'UT-1'}]),
            'tblutovarstavke':table('tblUtovarStavke',['UtovarID'],[{'UtovarID':'UT-1'}]),
        }
        self.scope={'format':'MIG01.3-scope-proposal','source_sha256':'abc',
                    'last_real_row_inclusive':{'tblOtkup':1,'tblKooperanti':1,'tblZbirna':1,'tblNovac':1},
                    'all_test_tables':['tblUtovar'],
                    'review_rows':{'tblNovac':[3]},
                    'review_entire_tables':['tblUtovarStavke']}

    def test_noncontiguous_novac_review_and_unknown(self):
        states=parse_scope(self.scope,self.tabs,'abc')
        self.assertEqual(states['tblNovac'],[KEPT,TEST,REVIEW])
        self.assertEqual(states['tblUtovar'],[TEST])
        self.assertEqual(states['tblUtovarStavke'],[REVIEW])
        self.assertEqual(states['tblOtkup'],[KEPT,TEST])

    def test_integrity_links_kept_not_mapped_to_test(self):
        self.tabs['tblotkup'].rows[0]['KooperantID']='KOP-TEST'
        r=evaluate(self.tabs,self.scope,'abc')
        kinds={x['kind'] for x in r['exact_id_links']['examples']}
        self.assertIn('KEEP_CANDIDATE_TO_TEST_CANDIDATE',kinds)

    def test_split_generation_detected(self):
        r=evaluate(self.tabs,self.scope,'abc')
        self.assertEqual(len(r['document_group_boundary']),1)
        self.assertEqual(r['document_group_boundary'][0]['table'],'tblZbirna')
        self.assertEqual(r['document_group_boundary'][0]['ordinals'],[1,2])

    def test_kept_missing_zbirna_and_no_raw_values(self):
        r=evaluate(self.tabs,self.scope,'abc')
        self.assertEqual(r['zbirna_number_links']['tblOtkup']['counts']['NOT_FOUND_ANYWHERE'],1)
        dumped=str(r)
        self.assertNotIn('ZBR-MISSING',dumped)
        self.assertNotIn('KOP-TEST',dumped)

    def test_invalid_source_sha_rejected(self):
        with self.assertRaises(AuditError):parse_scope(self.scope,self.tabs,'mismatch')
        with self.assertRaises(AuditError):
            altered={**self.scope,'last_real_row_inclusive':{'tblOtkup':5}}
            parse_scope(altered,self.tabs,'abc')

    def test_unknown_table_raises_and_no_silent_omission(self):
        bad={**self.scope,'last_real_row_inclusive':{'tblNonexistent':0}}
        with self.assertRaises(AuditError):parse_scope(bad,self.tabs,'abc')

if __name__=='__main__':unittest.main()
