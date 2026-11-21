"""Synthetic tests for targeted historical lineage, no real personal data."""
import json
import unittest
from legacy_audit import Table, AuditError
from legacy_coldchain_lineage import analyze,human,successor_evidence


def table(name,rows):
    return Table(name,name,list(dict.fromkeys(key for row in rows for key in row)),rows)


class LineageTests(unittest.TestCase):
    def setUp(self):
        ts=[
            table('tblPrijemnica',[
                {'PrijemnicaID':'PRJ-OLD','Stornirano':'DA','KupacID':'K-1','Klasa':'I','ZamenjenSa':'PRJ-NEW'},
                {'PrijemnicaID':'PRJ-NEW','Stornirano':'NE','KupacID':'K-1','Klasa':'I','IspravkaOd':'PRJ-OLD'},
                {'PrijemnicaID':'PRJ-TST','Stornirano':'NE','KupacID':'K-2','Klasa':'I'},
                {'PrijemnicaID':'PRJ-OTHER','Stornirano':'DA','KupacID':'K-2','Klasa':'II'}]),
            table('tblPaleta',[
                {'PaletaID':'PAL-1','Stornirano':'NE'},
                {'PaletaID':'PAL-2','Stornirano':'NE'},
                {'PaletaID':'PAL-TST','Stornirano':'DA','BrojPalete':'TST-PRO-123'}]),
            table('tblPaletaStavka',[
                {'StavkaID':'PS-1','PaletaID':'PAL-1','PrijemnicaID':'PRJ-OLD','Stornirano':'NE'},
                {'StavkaID':'PS-2','PaletaID':'PAL-1','PrijemnicaID':'PRJ-TST','Stornirano':'NE'},
                {'StavkaID':'PS-3','PaletaID':'PAL-2','PrijemnicaID':'PRJ-OTHER','Stornirano':'NE'},
                {'StavkaID':'PS-4','PaletaID':'PAL-TST','PrijemnicaID':'PRJ-NEW','Stornirano':'DA'}]),
            table('tblPrerada',[{'PreradaID':'PRE-1','Stornirano':'NE'},
                                {'PreradaID':'PRE-2','Stornirano':'DA'}]),
            table('tblPreradaStavka',[
                {'StavkaID':'PST-1','PreradaID':'PRE-1','PaletaID':'PAL-1','Stornirano':'NE'},
                {'StavkaID':'PST-2','PreradaID':'PRE-2','PaletaID':'PAL-TST','Stornirano':'DA'}])]
        self.tables={t.name.lower():t for t in ts}
        self.scope={'format':'MIG01.3-scope-proposal','source_sha256':'SYNTHETIC',
                    'last_real_row_inclusive':{'tblPrijemnica':2},'all_test_tables':[],
                    'review_rows':{},'review_entire_tables':[]}
        self.report=analyze(self.tables,self.scope,'SYNTHETIC')

    def test_review_just_two_active_links_and_correct_unique_counts(self):
        x=self.report['active_pallet_items_to_stornirana_receipts']
        self.assertEqual(x['count'],2)
        self.assertEqual(x['distinct_receipts'],2)
        self.assertEqual(x['distinct_pallets'],2)
        self.assertEqual(x['distinct_processing_headers'],1)

    def test_explicit_successor_evidence_not_action(self):
        x=self.report['active_pallet_items_to_stornirana_receipts']['cases'][0]
        self.assertEqual(x['successor_evidence'],'UNIQUE_EXPLICIT_ACTIVE_CORE_SUCCESSOR_CANDIDATE')
        self.assertEqual(x['explicit_successor_candidate_row'],2)
        self.assertFalse(self.report['migration_authorized'])

    def test_mixed_pallet_test_context_surfaces_to_processing(self):
        a=self.report['active_pallet_items_to_stornirana_receipts']['cases'][0]
        self.assertEqual(a['test_receipt_items_on_same_pallet'],1)
        flags=self.report['processing_graph_review']['flag_counts']
        self.assertIn({'flag':'SOURCE_PALLET_MIXES_CORE_TEST','state':'ACTIVE','count':1},flags)
        self.assertIn({'flag':'SOURCE_PALLET_ACTIVE_ITEM_ON_STORNIRANA_RECEIPT','state':'ACTIVE','count':1},flags)

    def test_receipt_scope_difference(self):
        by=self.report['active_pallet_items_to_stornirana_receipts']['by_receipt_scope']
        self.assertEqual(by['CORE_CANDIDATE'],1)
        self.assertEqual(by['TEST_CANDIDATE'],1)

    def test_guard_disallows_cross_owner_automatic_target(self):
        ts=self.tables['tblprijemnica'];by={'PRJ-OLD':[1],'PRJ-NEW':[2]}
        ts.rows[1]['KupacID']='K-WRONG'
        e=successor_evidence(1,ts,by,['KEEP_CANDIDATE']*4)
        self.assertEqual(e['status'],'EXPLICIT_SUCCESSOR_OUT_OF_SCOPE_OR_MISMATCH')

    def test_privacy_raw_values_not_exported(self):
        txt=human(self.report)+json.dumps(self.report)
        for raw in ('PRJ-OLD','PRJ-NEW','PRJ-TST','PAL-1','PRE-1','K-1','TST-PRO'):
            self.assertNotIn(raw,txt)

    def test_wrong_source_sha_fails_closed(self):
        with self.assertRaises(AuditError):
            analyze(self.tables,self.scope,'OTHER')

    def test_missing_table_stops(self):
        x=dict(self.tables);del x['tblpreradastavka']
        with self.assertRaises(AuditError):analyze(x,self.scope,'SYNTHETIC')


if __name__=='__main__':unittest.main()
