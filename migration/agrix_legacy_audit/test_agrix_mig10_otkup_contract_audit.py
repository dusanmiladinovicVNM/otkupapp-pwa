import unittest
import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from agrix_mig10_otkup_contract_audit import scan, num, is_same, EXPECTED

class TestAudit(unittest.TestCase):
    def rows(self,overrides=None,so=None):
        h=[];l=[];s=[]
        for i in range(EXPECTED):
            oid='OTK-TEST-%05d'%i
            o={'OtkupID':oid,'Klasa':'I','Kolicina':'100','Cena':'120','KolAmbalaze':'5','BrutoKg':'105'}
            line={**o,'RedniBroj':'1'};header={'OtkupID':oid}
            h.append(header);l.append(line);s.append(dict(o))
        if overrides: l[122].update(overrides)
        if so:s[122].update(so)
        return h,l,s
    def test_clean(self):
        r=scan(*self.rows())
        self.assertEqual(r['status'],'OTKUP_CONTRACT_PASS')
        self.assertEqual(r['affected_source_row_count'],0)
    def test_historical_zero_price(self):
        r=scan(*self.rows({'Cena':'0'},{'Cena':'0'}),'OTK-TEST-00122')
        self.assertEqual(r['issue_counts']['PRICE_NONPOSITIVE'],1)
        self.assertEqual(r['origin_of_flagged_rows']['unchanged_historical_source'],1)
        self.assertEqual(r['lookup_source_row_ordinals'],[123])
        self.assertNotIn('OTK-TEST-00122',str(r))
    def test_stage_delta(self):
        r=scan(*self.rows({'Cena':'0'}))
        self.assertEqual(r['issue_counts']['SOURCE_TO_STAGE_MISMATCH_CENA'],1)
        self.assertEqual(r['origin_of_flagged_rows']['stage_differs_from_source'],1)
    def test_quantity_negative_and_bruto(self):
        r=scan(*self.rows({'Kolicina':'-1','BrutoKg':'-5','KolAmbalaze':'-2'},{'Kolicina':'-1','BrutoKg':'-5','KolAmbalaze':'-2'}))
        self.assertEqual(r['issue_counts']['QUANTITY_NONPOSITIVE'],1)
        self.assertEqual(r['issue_counts']['BRUTO_NEGATIVE'],1)
        self.assertEqual(r['issue_counts']['AMBALAZA_NEGATIVE'],1)
    def test_non_number(self):
        r=scan(*self.rows({'Cena':'NaN'},{'Cena':'NaN'}))
        self.assertEqual(r['issue_counts']['PRICE_NOT_NUMERIC_OR_MISSING'],1)
        self.assertIsNone(num('Infinity'))
    def test_invalid_class(self):
        r=scan(*self.rows({'Klasa':'III'},{'Klasa':'III'}))
        self.assertEqual(r['issue_counts']['INVALID_CLASS'],1)
    def test_bad_amount_types(self):
        r=scan(*self.rows({'KolAmbalaze':'2.5','BrutoKg':'1'},{'KolAmbalaze':'2.5','BrutoKg':'1'}))
        self.assertEqual(r['issue_counts']['AMBALAZA_NOT_INTEGER'],1)
        self.assertEqual(r['issue_counts']['BRUTO_BELOW_NET'],1)
    def test_empty_valid_packaging(self):
        h,l,s=self.rows({'KolAmbalaze':'','BrutoKg':''},{'KolAmbalaze':'','BrutoKg':''})
        self.assertEqual(scan(h,l,s)['affected_source_row_count'],0)
    def test_source_mismatch_guard(self):
        h,l,s=self.rows();l[122]['OtkupID']='wrong'
        r=scan(h,l,s)
        self.assertEqual(r['issue_counts']['PARENT_SOURCE_LINE_ID_MISMATCH'],1)
    def test_truncated_refused(self):
        h,l,s=self.rows()
        with self.assertRaises(Exception):scan(h,l[:1],s)
    def test_no_exposure(self):
        r=scan(*self.rows({'Cena':'0'},{'Cena':'0'}))
        self.assertNotIn('OtkupID=',str(r))
        self.assertTrue(r['customer_ids_names_prices_quantities_dates_not_exported'])

if __name__=='__main__': unittest.main()
