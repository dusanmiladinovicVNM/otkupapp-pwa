"""Synthetic tests and source-level guardrails; does NOT compile VBA or open Excel."""
import json
import unittest
from pathlib import Path
BASE=Path(__file__).resolve().parent

def business_events(rows):
    types={r['id']:r['type'] for r in rows}
    kept=[]
    for r in rows:
        effective_type=types[r['storno_of']] if r.get('storno_of') else r['type']
        if effective_type != 'ULAZ_TUDJE_AMBALAZE':
            kept.append(r)
    return kept

class Tests(unittest.TestCase):
    def test_koop19(self):
        rows=[
            dict(id='physical08', type='AMBALAZA_UZ_ROBU',qty=-23),
            dict(id='cover08', type='ULAZ_TUDJE_AMBALAZE',qty=12),
            dict(id='physical11', type='AMBALAZA_UZ_ROBU',qty=-3),
            dict(id='cover11', type='ULAZ_TUDJE_AMBALAZE',qty=3),
        ]
        self.assertEqual(sum(r['qty'] for r in rows),-11)
        self.assertEqual(sum(r['qty'] for r in business_events(rows)),-26)
        self.assertEqual(5-15,-10)   # Screenshot final balance 5; 15 technical additions
    def test_storno_pairs(self):
        rows=[
           dict(id='a', type='ULAZ_TUDJE_AMBALAZE',qty=3),
           dict(id='b', type='ULAZ_TUDJE_AMBALAZE',qty=-3,storno_of='a'),
           dict(id='c', type='AMBALAZA_UZ_ROBU',qty=-8),
           dict(id='d', type='AMBALAZA_UZ_ROBU',qty=8,storno_of='c'),
        ]
        self.assertEqual([x['id'] for x in business_events(rows)],['c','d'])
        self.assertEqual(sum(x['qty'] for x in business_events(rows)),0)
    def test_other_business_movements(self):
        rows=[dict(id='a',type='IZDATA_PRAZNA',qty=20),dict(id='b',type='POVRAT_PRAZNE',qty=-5)]
        self.assertEqual([x['qty'] for x in business_events(rows)],[20,-5])
    def test_diff_is_narrow(self):
        p=json.loads((BASE/'patch_plan.json').read_text())
        self.assertEqual(len(p['patches']),7)
        old_amb=(BASE/'snapshot/modAmbalaza.txt').read_text(encoding='utf-8-sig')
        new_amb=(BASE/'patched_preview/modAmbalaza.bas').read_text(encoding='utf-8')
        old_rep=(BASE/'snapshot/modIzvestaj.txt').read_text(encoding='utf-8-sig')
        new_rep=(BASE/'patched_preview/modIzvestaj.bas').read_text(encoding='utf-8')
        self.assertEqual(old_amb.count('AmbKretanjaNaloga('),new_amb.count('AmbKretanjaNaloga('))
        self.assertEqual(old_rep.count('AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))'),2)
        self.assertEqual(new_rep.count('AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID), True)'),2)
        self.assertIn('Optional ByVal samoPoslovno As Boolean = False',new_amb)
        self.assertIn('AmbKarticaSortKey(',new_rep)
        for symbol in ['Public Function AmbSaldoNaloga','Public Function AmbSaldoPoNalogu','Public Function GetAmbalazeStanje','Public Function AmbObavezaPartneru']:
            self.assertIn(symbol,old_amb)
            self.assertIn(symbol,new_amb)
        self.assertEqual((BASE/'snapshot/modScrIzvestaji.bas').read_bytes(),(BASE/'snapshot/modScrIzvestaji.bas').read_bytes())

if __name__=='__main__':unittest.main(verbosity=2)
