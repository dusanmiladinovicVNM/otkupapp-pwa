import unittest
import sys
from pathlib import Path
from unittest import mock
sys.path.insert(0, str(Path(__file__).resolve().parent))
import agrix_mig102_otkup_read_policy as m

SRC = '''Public Function VrednostOtkupa(ByVal otkupID As String) As Double
    If kol <= 0 Or cena <= 0 Then Err.Raise 123
End Function

Public Function StavkeOtkupaRedovi() As Variant
    RequireJedinstvenZaglavljeOtkupa zagl, SRC

    Dim imaStavku As Object
    For i = 1 To UBound(d, 1)
            RequireStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _
                                d(i, cAmb), i, SRC
    Next i
End Function

Private Function ZaglavljaOtkupaPoID(ByVal sourceName As String) As Object
End Function

' JEDAN LOGICKI DOKUMENT = JEDAN ID (A1/S4). Dva zaglavlja sa istim OtkupID
Private Sub RequireStavkaUgovor(ByVal zagl As Object, ByVal oid As String, _
                                ByVal klasa As Variant, ByVal kol As Variant, _
                                ByVal cena As Variant, ByVal amb As Variant, _
                                ByVal red As Long, ByVal sourceName As String)
    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then
        Err.Raise vbObjectError + 1906, sourceName, _
              "Kolicina i cena"
    End If
End Sub
Private Sub CreateOtkup_TX()
End Sub
'''

class TestReadPolicy(unittest.TestCase):
    def test_patch_narrow(self):
        patched=m.patch_module(SRC)
        self.assertIn('Set storniranaZaglavlja = OtkupStorniranaZaglavlja()',patched)
        self.assertIn('(CDbl(kol) = 0 And Not storniranaZaglavlja.Exists(oid))',patched)
        self.assertIn('Private Function OtkupStorniranaZaglavlja()',patched)
        self.assertIn('UCase$(Trim$(NzToText(d(i, cSt)))) = "DA"',patched)
        self.assertEqual(SRC.split('Public Function VrednostOtkupa(',1)[1].split('End Function',1)[0],
                         patched.split('Public Function VrednostOtkupa(',1)[1].split('End Function',1)[0])
        self.assertEqual(SRC.split('Private Sub CreateOtkup_TX',1)[1],
                         patched.split('Private Sub CreateOtkup_TX',1)[1])
    def test_double_patch_fails(self):
        with self.assertRaises(m.AuditError): m.patch_module(m.patch_module(SRC))
    def test_missing_anchor_fails(self):
        with self.assertRaises(m.AuditError): m.patch_module(SRC.replace('    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then',''))
    def test_extra_anchor_fails(self):
        with self.assertRaises(m.AuditError): m.patch_module(SRC + '    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then\n        Err.Raise vbObjectError + 1906, sourceName, _\n')
    def test_helper_provenance_policy(self):
        # Model the exact boolean formula (not a VBA COM execution).
        def accepted(status,q,p):
            st=(status.strip().upper() == 'DA')
            return p>0 and q>=0 and (q>0 or st)
        self.assertTrue(accepted('Da',0,1))
        self.assertFalse(accepted('',0,1))
        self.assertFalse(accepted('Ne',0,1))
        self.assertFalse(accepted('Da',-1,1))
        self.assertFalse(accepted('Da',0,0))
        self.assertFalse(accepted('Da',0,-1))
        self.assertTrue(accepted('',1,1))
        self.assertFalse(accepted('True',0,1))
    def test_precheck_requires_native_ui_status(self):
        rows=[{'OtkupID':f'k{i}', 'Stornirano':'Ne','Kolicina':'1','Cena':'2','KolAmbalaze':'0'} for i in range(m.EXPECTED)]
        lines=[dict(v) for v in rows]
        rows[122].update({'Stornirano':'1','Kolicina':'0','KolAmbalaze':'1'})
        lines[122].update({'Stornirano':'1','Kolicina':'0','KolAmbalaze':'1'})
        result={'issue_counts':{'QUANTITY_NONPOSITIVE':1},'affected_source_row_count':1,
                'issues_by_source_row_ordinal':[{'source_row_ordinal':123}], 'stage_rows':m.EXPECTED}
        with mock.patch.object(m,'git',return_value=m.EXPECTED_COMMIT), \
             mock.patch.object(m,'audit',return_value=result), \
             mock.patch.object(m,'load_jsonl',side_effect=[rows,rows,lines]), \
             mock.patch.object(m.hashlib,'sha256') as hs, \
             mock.patch.object(m.Path,'read_bytes',return_value=b'x'):
            hs.return_value.hexdigest.return_value=m.SCHEMA_SHA
            with self.assertRaisesRegex(m.AuditError,'Stornirano=Da'):
                m.precheck(Path('repo'),Path('stage'),123)

if __name__ == '__main__': unittest.main()
