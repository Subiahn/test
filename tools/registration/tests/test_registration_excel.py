import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook,load_workbook
from make_registration import convert
from registration_rules import read_excel

class ExportTests(unittest.TestCase):
    def test_conversion_manual_edit_registration(self):
        with tempfile.TemporaryDirectory() as t:
            src,out=Path(t)/'lookup.xlsx',Path(t)/'registration.xlsx'
            w=Workbook();s=w.active;s.append(['IP','Port','Application','Hostname','Protocol'])
            s.append(['192.0.2.1',8443,'TCP_OTHER','web01','TCP'])
            s.append(['192.0.2.2',443,'HTTPS','','TCP']);w.save(src);w.close()
            self.assertEqual(convert(src,out)['written'],1)
            w=load_workbook(out);s=w.active
            cols={c.value:c.column for c in s[1]}
            self.assertIsNone(s.cell(2,cols['ApplicationName']).value)
            self.assertIsNone(s.cell(2,cols['RegistrationType']).value)
            self.assertEqual(s.cell(2,cols['Hostname']).value,'web01')
            for k,v in {'Register':'Y','ApplicationName':'MANUAL_WEB','RegistrationType':'SBA','ParentApplication':'TCP'}.items():s.cell(2,cols[k],v)
            w.save(out);w.close()
            self.assertEqual(read_excel(out)[0]['name'],'MANUAL_WEB')
            with self.assertRaises(ValueError):convert(src,out)
if __name__=='__main__':unittest.main()
