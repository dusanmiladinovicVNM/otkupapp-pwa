import json
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile
from xml.sax.saxutils import escape

import agrix_mig19_5_evidence_and_parity as m
from agrix_mig19_3_parity import ledger_views, reconcile


def old(aid,station,day='2026-06-27',qty=300,smer='Ulaz',doc='OMUlaz'):
    return dict(AmbID=aid,Datum=day,TipAmbalaze='0,4kg',Kolicina=qty,Smer=smer,
      EntitetTip='Stanica',EntitetID=station,Stornirano='',DokumentTIP=doc,
      DokumentID='OTK-1',VozacID=station)


def target(aid,od,na,day='2026-06-27',qty=300,doc='AmbalazaDokument'):
    return dict(AmbID=aid,Datum=day,TipAmbalaze='0.4kg',Kolicina=qty,
      OdNalogTip=od[0],OdNalogID=od[1],NaNalogTip=na[0],NaNalogID=na[1],
      VrstaKretanja='PRENOS_INTERNO',DokumentTIP=doc,DokumentID='ADK-1',StornoOd='')


def build_test_xlsm(path, row):
    headers=['AmbID','Datum','TipAmbalaze','Kolicina','OdNalogTip','OdNalogID',
             'NaNalogTip','NaNalogID','VrstaKretanja','StornoOd','DokumentTIP','DokumentID']
    def excelcol(i):return chr(ord('A')+i)
    def cells(vals,r):return ''.join(
        f'<c r="{excelcol(i)}{r}" t="inlineStr"><is><t>{escape(str(v))}</t></is></c>'
        for i,v in enumerate(vals))
    root='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    nsrel='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg='http://schemas.openxmlformats.org/package/2006/relationships'
    wb=f'<workbook xmlns="{root}" xmlns:r="{nsrel}"><sheets><sheet name="Ambalaza" sheetId="1" r:id="rId1"/></sheets></workbook>'
    wbr=f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{nsrel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    sr=f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{nsrel}/table" Target="../tables/table1.xml"/></Relationships>'
    sh=f'<worksheet xmlns="{root}"><sheetData><row r="1">{cells(headers,1)}</row><row r="2">{cells([row.get(x,"") for x in headers],2)}</row></sheetData></worksheet>'
    table=f'<table xmlns="{root}" name="tblAmbalaza" displayName="tblAmbalaza" id="1" ref="A1:L2" />'
    with ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml',wb)
        z.writestr('xl/_rels/workbook.xml.rels',wbr)
        z.writestr('xl/worksheets/sheet1.xml',sh)
        z.writestr('xl/worksheets/_rels/sheet1.xml.rels',sr)
        z.writestr('xl/tables/table1.xml',table)


class Tests(unittest.TestCase):
    def test_old_signed_allows_negative(self):
        self.assertEqual(m.old_signed(old('A','ST-01',qty=500,smer='Izlaz')),Decimal(-500))
    def test_source_six_signature_and_unproved_lineage(self):
        source=[old('AMB-02317','ST-00005')]
        n=[target('NEW-1',('STANICA','ST-00001'),('STANICA','ST-00005'))]
        six,q700,_=m.known_cases(source,n,[])
        r=next(x for x in six if x['Case']=='AMB-02317')
        self.assertTrue(r['SourceSignatureVerified'])
        self.assertFalse(r['ExactLineageProven'])
        self.assertEqual(r['CanonicalSameDestinationCandidates'],1)
    def test_q700_requires_actual_source_evidence(self):
        source=[old('A-'+str(i),'ST-00001') for i in range(699)]
        source.append(old('AMB-00700','ST-00004',day='2026-06-22',qty=500,smer='Izlaz',doc='Otpremnica'))
        _,q,found=m.known_cases(source,[],[dict(NalogID='ST-00004',TipAmbalaze='0,4KG',Datum='2026-06-22',PromenaRazlike='500')])
        self.assertTrue(found);self.assertTrue(q['MatchesST04Minus500Hypothesis'])
        source[699]['Kolicina']=499
        _,q,_=m.known_cases(source,[],[])
        self.assertFalse(q['MatchesST04Minus500Hypothesis'])
    def test_breakdown_recomputes_exact_station_effect(self):
        a=old('AMB-1','ST-00002',day='2026-06-29',qty=83)
        changes=[dict(NalogTip='STANICA',NalogID='ST-00002',TipAmbalaze='0,4KG',
                      Datum='2026-06-29',StariDnevniEfekat='83',NoviDnevniEfekat='0',
                      PromenaRazlike='-83',RazlikaOdPocetka='-83')]
        oldrows,newrows,diagnostic=m.observe_source_and_target([a],[],changes)
        self.assertEqual((len(oldrows),len(newrows)),(1,0))
        self.assertEqual(diagnostic[0]['LegacyPositiveRows'],1)
    def test_doc_balancing_not_interpreted_as_pass(self):
        idx=('0,4KG','2026-06-01','OTKUP','OTK-1')
        a=('KOOPERANT','KO-01',*idx)
        b=('STANICA','ST-01',*idx)
        olddoc={a:Decimal(10),b:Decimal(-10)}
        newdoc={a:Decimal(-10),b:Decimal(10)}
        priv,cls,summary=m.docs_by_key(olddoc,newdoc)
        self.assertEqual(summary['documents_with_otkup_delta'],1)
        self.assertEqual(summary['document_groups_whose_participant_deltas_sum_zero'],1)
        self.assertIn('NOT per-document account parity',summary['warning'])
    def test_naive_633_debit_breaks_st1(self):
        gaps=[dict(NalogID=st,TipAmbalaze='0,4KG',Razlika=gap)
              for st,gap in [('ST-00001','-18'),('ST-00002','-83'),('ST-00004','500'),('ST-00005','-550')]]
        sim=m.synthetic_naive_projection(gaps)
        self.assertEqual(sim['BlindSixMovesAndThreeInternalTransfers_ST1Delta'],'-633')
        self.assertEqual(sim['ST2Delta'],'0');self.assertEqual(sim['ST5Delta'],'0')
    def test_run_synthetic_readonly_minimal_ooxml(self):
        with tempfile.TemporaryDirectory() as temp:
            root=Path(temp);stage=root/'stage';stage.mkdir();out=root/'results'
            src=stage/'UNCONVERTED__tblAmbalaza.jsonl'
            oldrow=old('OLD','ST-00002',qty=5,day='2026-06-01')
            src.write_text(json.dumps(oldrow)+'\n')
            dest=root/'QA.xlsm'
            newrow=target('NEW',('STANICA','ST-00001'),('STANICA','ST-00002'),day='2026-06-01',qty=5)
            build_test_xlsm(dest,newrow)
            oldstage=m.STAGE_SHA
            try:
                m.STAGE_SHA=m.sha256(src)
                source_before=m.sha256(src);target_before=m.sha256(dest)
                exit_code=m.main(['--stage',str(stage),'--workbook',str(dest),
                   '--out-dir',str(out),'--expected-target-sha',target_before])
                self.assertEqual(exit_code,0)
                self.assertTrue((out/'MIG195_SUMMARY.json').is_file())
                self.assertEqual(m.sha256(src),source_before)
                self.assertEqual(m.sha256(dest),target_before)
            finally:m.STAGE_SHA=oldstage

if __name__=='__main__':unittest.main()
