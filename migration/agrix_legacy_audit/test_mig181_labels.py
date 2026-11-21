import unittest
from agrix_mig181_qa_amb_labels import patch_izvestaj, patch_ambalaza

AMB = '''Attribute VB_Name = "modAmbalaza"
' Vraca 2D niz (1 To n, 1 To 8):
'   7 = ProtivpartnerTip   8 = ProtivpartnerID
Public Function AmbKretanjaNaloga(ByVal tip As String, ByVal id As String) As Variant
    Dim cStorno As Long, cAmbID As Long
    cStorno = kolIdx(COL_AMB_STORNO_OD)
    cAmbID = kolIdx(COL_AMB_ID)
    ReDim res(1 To poz.count, 1 To 8)
        res(n, 6) = otkazani.Exists(AmbText(data(ri, cAmbID)))
End Function
''' 
IZV='''Attribute VB_Name = "modIzvestaj"
Public Function ReportKarticaAmbalaze(ByVal kooperantID As String, _
                                       ByVal datumOd As Date, _
                                       ByVal datumDo As Date) As Variant
    ambData = modAmbalaza.AmbKretanjaNaloga(AMB_NALOG_KOOPERANT, Trim$(kooperantID))
    If Not IsArray(ambData) Then
        ReportKarticaAmbalaze = Empty
        Exit Function
    End If
                    Dim opis As String
                    opis = NzToText(ambData(i, colTip))   ' TipAmbalaze
                    If lbl <> "" Then opis = Trim$(opis & " (" & lbl & ")")

                    moves.Add Array(d, brojDok, opis, ulaz, izlaz)
    arr = SortArray(tmpArr, 6, True)
End Function
Private Function AmbKarticaSortKey(ByVal d As Date, ByVal broj As String, _
                                   ByVal ulaz As Double, ByVal ordinal As Long) As String
End Function
Private Function BuildOtkupBrojDokDict() As Object
End Function
'''

class Tests(unittest.TestCase):
    def test_patch_amb_contract(self):
        out=patch_ambalaza(AMB)
        self.assertIn("' Vraca 2D niz (1 To n, 1 To 10)",out)
        self.assertIn('cVrstaKretanja = kolIdx(COL_AMB_VRSTA_KRETANJA)',out)
        self.assertIn('res(n, 9) = AmbText(data(ri, cVrstaKretanja))',out)
        self.assertIn('res(n, 10) = AmbText(data(ri, cStorno))',out)
        self.assertIn('res(n, 6) = otkazani.Exists',out)
        self.assertEqual(out.count('Public Function AmbKretanjaNaloga'),1)
        with self.assertRaisesRegex(ValueError,'anchor'):
            patch_ambalaza(out)

    def test_labels_map_and_preserve_sort(self):
        out=patch_izvestaj(IZV)
        for key in ['AMB_VK_IZDATA_PRAZNA','AMB_VK_UZ_ROBU','AMB_VK_ULAZ_TUDJE','AMB_VK_OTPIS']:
            self.assertIn('Case '+key, out)
        self.assertIn('AmbKarticaNazivKretanja(NzToText(ambData(i, 9)))',out)
        self.assertIn('If UBound(ambData, 2) < 10 Then',out)
        self.assertIn('arr = SortArray(tmpArr, 6, True)',out)
        self.assertIn('Private Function AmbKarticaSortKey(',out)
        self.assertEqual(out.count('Private Function AmbKarticaNazivKretanja('),1)
        self.assertIn('storno: ',out)
        with self.assertRaisesRegex(ValueError,'already installed'):
            patch_izvestaj(out)

    def test_no_base_report(self):
        with self.assertRaisesRegex(ValueError,'NOT the MIG-18'):
            patch_izvestaj(IZV.replace('SortArray(tmpArr, 6, True)','SortArray(tmpArr, 1, True, 2)'))

if __name__=='__main__':unittest.main()
