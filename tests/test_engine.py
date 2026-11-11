import tempfile, unittest
from pathlib import Path
from docx import Document
from openpyxl import Workbook, load_workbook
from engine import detect_placeholders, replace_placeholders, load_sheet, safe_filename

class EngineTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.template = self.root / 'template.docx'
        doc = Document()
        p = doc.add_paragraph('Student: ')
        p.add_run('{{Full ').bold = True
        p.add_run('Name}}')
        p2 = doc.add_paragraph('Course: {{Course}}')
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = 'Email'
        table.cell(0, 1).text = '{{Email Address}}'
        nested = table.cell(0, 0).add_table(rows=1, cols=1)
        nested.cell(0, 0).text = '{{Section}}'
        doc.sections[0].header.paragraphs[0].text = 'Header {{Full Name}}'
        doc.sections[0].footer.paragraphs[0].text = 'Footer {{Course}}'
        doc.save(self.template)
    def tearDown(self): self.tmp.cleanup()
    def test_detect(self):
        self.assertEqual(set(detect_placeholders(self.template)), {'Full Name','Course','Email Address','Section'})
    def test_replace_paragraph_table_header_footer_split_runs(self):
        out = self.root / 'out.docx'
        mapping = {'Full Name':'Full Name','Course':'Course','Email Address':'Email Address','Section':'Section'}
        result = replace_placeholders(self.template, out, {'Full Name':'Jane Doe','Course':'BSIT','Email Address':'jane@example.com','Section':'3A'}, mapping)
        self.assertEqual(result['unresolved_placeholders'], [])
        doc = Document(out)
        texts = [p.text for p in doc.paragraphs]
        self.assertTrue(any('Jane Doe' in t for t in texts))
        self.assertTrue(any('BSIT' in t for t in texts))
        self.assertEqual(doc.tables[0].cell(0,1).text, 'jane@example.com')
        self.assertIn('3A', doc.tables[0].cell(0,0).tables[0].cell(0,0).text)
        self.assertIn('Jane Doe', doc.sections[0].header.paragraphs[0].text)
        self.assertIn('BSIT', doc.sections[0].footer.paragraphs[0].text)
    def test_excel_read(self):
        x = self.root / 'data.xlsx'; wb = Workbook(); ws = wb.active; ws.title = 'Students'
        ws.append(['Full Name','Course','Contact Number']); ws.append(['Jane Doe','BSIT','09170000001']); ws.append([None,None,None]); wb.save(x)
        headers, rows, warnings = load_sheet(x, 'Students')
        self.assertEqual(headers, ['Full Name','Course','Contact Number'])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['Contact Number'], '09170000001')
    def test_safe_filename(self):
        self.assertEqual(safe_filename('Jane/Doe'), 'Jane_Doe')

if __name__ == '__main__': unittest.main()
