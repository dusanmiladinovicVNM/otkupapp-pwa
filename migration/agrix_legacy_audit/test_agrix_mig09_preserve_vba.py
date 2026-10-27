"""Synthetic OOXML tests only. No customer data and no Excel COM."""
import io
import json
from pathlib import Path
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET
from unittest.mock import patch
import agrix_mig09_preserve_vba as m
from legacy_audit import AuditError,parse_xlsm
import agrix_mig07_goods_copy as g

X='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'

class TestMIG09OOXML(unittest.TestCase):
 def test_styles_three_new_xfs(self):
  src=(f'<styleSheet xmlns="{X}"><fonts count="1"><font/></fonts>'
       '<fills count="1"><fill/></fills><borders count="1"><border/></borders>'
       '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
       '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
       '</styleSheet>').encode()
  b,ids=m.inject_styles(src)
  root=ET.fromstring(b)
  self.assertEqual(ids,{'text':1,'date':2,'datetime':3})
  self.assertEqual(len(root.findall('{'+X+'}cellXfs/{'+X+'}xf')),4)
  self.assertEqual(len(root.findall('{'+X+'}numFmts/{'+X+'}numFmt')),2)

 def test_xml_rejects_invalid_chars_and_escapes_literal(self):
  with self.assertRaises(AuditError):m.legal_xml_text('\x01','t','x',4)
  v=m.make_rows([('000011','a & < b',2.5)],['Ref','Naziv','Cena'],['text','text','number'],{'text':1,'date':2,'datetime':3},'tblTest')
  self.assertIn(b'000011',v)
  self.assertIn(b'a &amp; &lt; b',v)
  self.assertIn(b'<v>2.5</v>',v)

 def test_no_vba_edit_and_25_table_roundtrip(self):
  # Synthesize the minimal Excel package with all 25 targets + 22 non-target
  # tables and a deliberately large dummy VBA blob (not an actual runnable VBA).
  others=['tblOther'+str(i) for i in range(22)]
  names=list(g.TABLE_ORDER)+others
  schema={n:{'table':n,'sheet':'Sheet'+str(i+1),'columns':['ID','Datum','Naziv']} for i,n in enumerate(names)}
  tables={}
  xrel=f'xmlns:r="{R}"'
  sheets=[];wrels=[]
  with tempfile.TemporaryDirectory() as td:
   inp=Path(td)/'blank.xlsm';out=Path(td)/'copied.xlsm'
   with zipfile.ZipFile(inp,'w',compression=zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
    z.writestr('xl/vbaProject.bin', b'BINARY-VBA-CHECK'+(b'Z'*100000))
    z.writestr('xl/styles.xml',f'<styleSheet xmlns="{X}"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
    for i,n in enumerate(names,1):
     sheets.append(f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>')
     wrels.append(f'<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>')
     sheetxml=(f'<worksheet xmlns="{X}" {xrel}><dimension ref="A1:C2"/><sheetData>'
          '<row r="1">'+''.join(f'<c r="{c}1" t="inlineStr"><is><t>{h}</t></is></c>' for c,h in zip('ABC',['ID','Datum','Naziv']))+'</row>'
          '<row r="2"><c r="A2"/></row></sheetData>'
          '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
     z.writestr(f'xl/worksheets/sheet{i}.xml',sheetxml)
     z.writestr(f'xl/worksheets/_rels/sheet{i}.xml.rels',f'<Relationships xmlns="{P}"><Relationship Id="rId1" Type="{R}/table" Target="../tables/table{i}.xml"/></Relationships>')
     tablexml=(f'<table xmlns="{X}" id="{i}" name="{n}" displayName="{n}" ref="A1:C2" headerRowCount="1">'
      '<autoFilter ref="A1:C2"/><tableColumns count="3">'
      '<tableColumn id="1" name="ID"/><tableColumn id="2" name="Datum"/><tableColumn id="3" name="Naziv"/>'
      '</tableColumns><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>')
     z.writestr(f'xl/tables/table{i}.xml',tablexml)
    z.writestr('xl/workbook.xml',f'<workbook xmlns="{X}" {xrel}><sheets>'+''.join(sheets)+'</sheets></workbook>')
    z.writestr('xl/_rels/workbook.xml.rels',f'<Relationships xmlns="{P}">'+''.join(wrels)+'</Relationships>')
   stage={n:[(f'{n}-00001', 46232.5,'Text & <literal>')] for n in g.TABLE_ORDER}
   stage['tblStanice'].append(('00002',46233.0,'Živa & data'))
   kinds={n:['text','date','text'] for n in g.TABLE_ORDER}
   before=m.sha(inp)
   audit=m.build_ooxml_copy(inp,out,schema,stage,kinds)
   self.assertEqual(audit['xml_parts_changed'],51)
   self.assertTrue(audit['vba_binary_preserved_exactly'])
   self.assertEqual(m.sha(inp),before)
   with zipfile.ZipFile(inp) as source,zipfile.ZipFile(out) as target:
    self.assertEqual(source.read('xl/vbaProject.bin'),target.read('xl/vbaProject.bin'))
    self.assertEqual(source.read('xl/workbook.xml'),target.read('xl/workbook.xml'))
    self.assertEqual(source.read('xl/tables/table47.xml'),target.read('xl/tables/table47.xml'))
   parsed=parse_xlsm(out)
   # Same post-save data comparator used for the real MIG-02 staging.
   verified=g.verify_data_roundtrip(out,stage,{n:['ID','Datum','Naziv'] for n in g.TABLE_ORDER},kinds)
   self.assertEqual(sum(verified.values()),len(g.TABLE_ORDER)+1)
   self.assertEqual(len(parsed),47)
   self.assertEqual(len(parsed['tblstanice'].rows),2)
   self.assertEqual(parsed['tblstanice'].rows[1]['ID'],'00002')
   self.assertEqual(parsed['tblstanice'].rows[1]['Naziv'],'Živa & data')
   self.assertEqual(parsed['tblstanice'].rows[0]['Naziv'],'Text & <literal>')
   self.assertEqual(parsed['tblstanice'].rows[0]['Datum'],'46232.5')
   self.assertEqual(parsed['tblother20'].rows,[])

if __name__=='__main__':unittest.main()
