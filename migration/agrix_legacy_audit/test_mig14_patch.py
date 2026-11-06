import unittest
from agrix_mig14_patch_modotkup import patch
class Tests(unittest.TestCase):
    def setUp(self):
        self.s='''Attribute VB_Name = "modOtkup"
Public Function StavkeOtkupaRedovi() As Variant
    RequireJedinstvenZaglavljeOtkupa zagl, SRC
            RequireStavkaUgovor zagl, oid, d(i, cKl), d(i, cKol), d(i, cCena), _
                                d(i, cAmb), i, SRC
End Function
Private Sub RequireJedinstvenZaglavljeOtkupa(ByVal zagl As Object, _
 ByVal src As String)
End Sub
Private Sub RequireStavkaUgovor(ByVal zagl As Object, ByVal oid As String, _
                                ByVal klasa As Variant, ByVal kol As Variant, _
                                ByVal cena As Variant, ByVal amb As Variant, _
                                ByVal red As Long, ByVal sourceName As String)
    If Len(oid) = 0 Then
        Err.Raise 0
    End If
    If CDbl(kol) <= 0 Or CDbl(cena) <= 0 Then
        Err.Raise 0
    End If
End Sub
'''
    def test_exact_patch(self):
        o=patch(self.s)
        self.assertIn('StorniraniOtkupiZaBulkRead(SRC)',o)
        self.assertIn('istorijskiStornirani.Exists(oid)',o)
        self.assertIn('CDbl(kol) < 0 Or CDbl(cena) <= 0',o)
        self.assertIn('Optional ByVal dozvoliNuluZaStorno As Boolean = False',o)
        self.assertIn('Attribute VB_Name = "modOtkup"',o)
        self.assertEqual(o.count('Private Function StorniraniOtkupiZaBulkRead('),1)
        with self.assertRaises(ValueError): patch(o)
    def test_fail_closed_if_anchor_differs(self):
        with self.assertRaises(ValueError): patch(self.s.replace('    RequireJedinstvenZaglavljeOtkupa zagl, SRC','    OtherCheck zagl, SRC'))
    def test_validator_logic(self):
        def rejected(k,c,cancel):return k<0 or c<=0 or (k==0 and not cancel)
        self.assertFalse(rejected(0,15,True))
        self.assertTrue(rejected(0,15,False))
        self.assertTrue(rejected(-1,15,True))
        self.assertTrue(rejected(0,0,True))
        self.assertFalse(rejected(1,15,False))
if __name__=='__main__':unittest.main()
