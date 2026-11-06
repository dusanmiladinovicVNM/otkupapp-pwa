"""Synthetic tests only. No customer workbook data."""
import unittest
from legacy_audit import Table,AuditError
from legacy_core_scope import evaluate,tiers


def table(name,rows):
    cols=list(dict.fromkeys(k for row in rows for k in row))
    return Table(name,name,cols,rows)

class TestCoreScope(unittest.TestCase):
    def setUp(self):
        self.tabs={t.name.lower():t for t in [
            table('tblOtkup',[{'OtkupID':'OT-1','BrojDokumenta':'1','StanicaID':'STA-1','Datum':'D','KooperantID':'K-1','KulturaID':'CUL-1','BrojZbirne':'NOMATCH'},
                              {'OtkupID':'OT-2','BrojDokumenta':'1','StanicaID':'STA-1','Datum':'D','KooperantID':'K-1','KulturaID':'CUL-1','BrojZbirne':'NOMATCH'}]),
            table('tblOtpremnica',[{'OtpremnicaID':'OP-1','GeneracijaID':'G-1','Klasa':'I','BrojOtpremnice':'1'},
                                   {'OtpremnicaID':'OP-2','GeneracijaID':'G-1','Klasa':'II','BrojOtpremnice':'1'}]),
            table('tblZbirna',[{'ZbirnaID':'Z-1','GeneracijaID':'G-2','BrojZbirne':'1','Klasa':'I'},
                               {'ZbirnaID':'Z-2','GeneracijaID':'G-2','BrojZbirne':'1','Klasa':'II'}]),
            table('tblPrijemnica',[{'PrijemnicaID':'P-1','FakturaID':'FAK-T1','BrojZbirne':'1'}]),
            table('tblFakture',[{'FakturaID':'FAK-T1','KupacID':'KUP-1'}]),
            table('tblFakturaStavke',[{'FakturaStavkaID':'FS-1','UtovarID':'UT-1'}]),
            table('tblUtovar',[{'UtovarID':'UT-1'}]),
            table('tblMagacin',[{'MagacinID':'MAG-1','ArtikalID':'ART-POC-DUG','Tip':'Izlaz','KooperantID':'K-1'}]),
            table('tblArtikli',[{'ArtikalID':'ART-POC-DUG'}]),
            table('tblTipAmbalaze',[{'TipAmbalaze':'12/1'}]),
            table('tblKooperanti',[{'KooperantID':'K-1'}]),
            table('tblStanice',[{'StanicaID':'STA-1'}]),
            table('tblKupci',[{'KupacID':'KUP-1'}]),
            table('tblVozaci',[{'VozacID':'V-1'}]),
            table('tblKulture',[{'KulturaID':'CUL-1'}]),
            table('tblParcele',[{'ParcelaID':'PAR-1'}]),
            table('tblAmbalaza',[{'AmbID':'A1'}]),
            table('tblNovac',[{'NovacID':'NOV-1','OtkupID':'OT-1'}]),
            table('tblBankaImport',[{'BankaImportID':'BI-1'}]),
            table('tblStornoVeze',[{'StornoVezaID':'SV-1'}]),
            table('tblStornoZurnal',[{'StornoZurnalID':'SZ-1'}]),
        ]}
        self.scope={
            'format':'MIG01.3-scope-proposal','source_sha256':'HASH',
            'last_real_row_inclusive': {
                'tblOtkup':1,'tblOtpremnica':1,'tblZbirna':1,'tblPrijemnica':1,
                'tblAmbalaza':1,'tblMagacin':1,'tblKooperanti':1,'tblStanice':1,
                'tblKupci':1,'tblVozaci':1,'tblKulture':1,'tblParcele':1,'tblNovac':1,
            },'all_test_tables':['tblUtovar'], 'review_rows':{},'review_entire_tables':[]}

    def test_fakture_are_out_of_scope_and_terminal_fk_does_not_block(self):
        r=evaluate(self.tabs,self.scope,'HASH')
        self.assertEqual(r['row_counts_by_role']['tblFakture'],{'OUT_OF_SCOPE':1})
        self.assertEqual(r['row_counts_by_role']['tblFakturaStavke'],{'OUT_OF_SCOPE':1})
        self.assertFalse(any(x['field']=='FakturaID' and x['source_table']=='tblPrijemnica'
                             for x in r['core_typed_fk_evidence']['by_type']))
        self.assertFalse(any(x['source_table']=='tblFakturaStavke' for x in r['core_typed_fk_evidence']['by_type']))

    def test_three_splits_are_reported_never_automatically_resolved(self):
        r=evaluate(self.tabs,self.scope,'HASH')
        self.assertEqual(r['outstanding']['cross_boundary_document_candidate_groups'],3)
        self.assertEqual(sum(x['basis']=='GEN' for x in r['boundary_documents']),2)

    def test_missing_zbirna_kept_remains_visible(self):
        r=evaluate(self.tabs,self.scope,'HASH')
        x=r['zbirna_number_not_found']
        self.assertEqual([(e['source_table'],e['source_ordinal']) for e in x],[('tblOtkup',1)])

    def test_agro_article_fk_validated_and_no_raw_values(self):
        r=evaluate(self.tabs,self.scope,'HASH')
        self.assertNotIn('FK_NOT_FOUND: tblMagacin.ArtikalID',str(r))
        self.assertNotIn('K-1',str(r))
        self.assertNotIn('NOMATCH',str(r))
        self.assertNotIn('FAK-T1',str(r))

    def test_finance_separate_from_core(self):
        r=evaluate(self.tabs,self.scope,'HASH')
        self.assertEqual(r['row_counts_by_role']['tblNovac'],{'FINANCE_CANDIDATE':1})
        self.assertEqual(r['row_counts_by_role']['tblBankaImport'],{'FINANCE_REVIEW':1})

    def test_hash_mismatch_fail_closed(self):
        with self.assertRaises(AuditError): tiers(self.tabs,self.scope,'WRONG')

if __name__=='__main__':unittest.main()
