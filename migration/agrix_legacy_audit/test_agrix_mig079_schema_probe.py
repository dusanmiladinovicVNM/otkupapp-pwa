import unittest
from types import SimpleNamespace
import agrix_mig079_schema_probe as p

class TestSchemaProbe(unittest.TestCase):
    def test_counts_canonical(self):
        self.assertEqual(sum(p.COUNTS), 9610)
        self.assertEqual(len(p.COUNTS),25)

    def test_synthetic_does_not_read_customer_values(self):
        schema={'tblStanice':['StanicaID','Naziv','Aktivan']}
        matrices,kinds=p.synthetic_matrices(schema,['tblStanice'])
        self.assertEqual(len(matrices['tblStanice']),8)
        self.assertEqual(matrices['tblStanice'][0][0], 'SYNTHETIC-MIG079-tblStanice-000001')
        self.assertEqual(matrices['tblStanice'][0][1:],('',''))

    def test_report_first_mismatch_is_public_schema_only(self):
        schema={'tblStanice':['StanicaID','Naziv','Mesto']}
        parsed={'tblstanice':SimpleNamespace(columns=['StanicaID','Naziv','Column3'],rows=[{'StanicaID':'SYNTHETIC-1'}])}
        r=p.column_report(['tblStanice'],schema,parsed)
        self.assertEqual(r['tblStanice']['status'],'MISMATCH')
        self.assertIn('first_different_position=3',r['tblStanice']['column_difference'])
        self.assertNotIn('SYNTHETIC-1',str(r))

    def test_xml_and_excel_header_are_both_checked(self):
        schema={'tblStanice':['StanicaID','Naziv']}
        parsed={'tblstanice':SimpleNamespace(columns=['StanicaID','Naziv'],rows=[])}
        com={'tblStanice':{'ListColumns':['StanicaID','Naziv'], 'HeaderRowRange':['StanicaID','HeaderChanged'], 'ListRows':1,'RangeRows':2}}
        r=p.column_report(['tblStanice'],schema,parsed,com)['tblStanice']
        self.assertEqual(r['status'],'MATCH')
        self.assertTrue(r['excel_list_columns_match'])
        self.assertFalse(r['excel_header_cells_match'])

if __name__=='__main__':unittest.main()
