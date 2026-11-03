import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook
from registration_rules import HEADERS, row_fields, read_excel, xml_for, parse_response, preflight, verify
import xml.etree.ElementTree as ET

class RegistrationTests(unittest.TestCase):
    def base(self, **kw):
        return dict(Register='Y',ApplicationName='ERP_WEB',RegistrationType='SBA',ParentApplication='TCP',Protocol='TCP',IP='192.0.2.1',Port='8443',**kw)
    def group(self,d=None):
        mode,f=row_fields(d or self.base()); return {'name':f['ApplicationShortName'],'mode':mode,'fields':f,'rows':[2]}
    def test_sba(self):
        g=self.group(); root=ET.fromstring(xml_for(g))
        self.assertEqual(root.findtext('.//ServerAddresses'),'192.0.2.1/32')
        self.assertEqual(root.findtext('.//ServerPortRange'),'8443')
        self.assertIsNone(root.find('.//ApplicationPort'))
    def test_child(self):
        d=self.base(); d.update(RegistrationType='CHILD',ParentApplication='HTTPS')
        g=self.group(d)
        self.assertEqual(g['fields']['HostAddresses'],'192.0.2.1/32')
        self.assertEqual(g['fields']['AdditionalPort'],'8443')
    def test_wka(self):
        d=self.base(); d.update(RegistrationType='WKA',IP='')
        self.assertEqual(self.group(d)['fields']['ApplicationPort'],'8443')
    def test_url(self):
        d=self.base(); d.update(RegistrationType='URL',ParentApplication='HTTPS',Port='',URLAddress='https://example.test',ExactMatch='true')
        self.assertEqual(self.group(d)['fields']['ApplicationType'],'URL Application')
    def test_invalid_options(self):
        for k,v in [('ApplicationName',''),('Port',''),('RegistrationType','MULTICAST'),('Logged','None')]:
            d=self.base();d[k]=v
            with self.subTest(k=k), self.assertRaises(ValueError): row_fields(d)
    def test_clients(self):
        d=self.base(ClientAddresses='192.0.2.0/24',ClientPortRange='1000-2000')
        with self.assertRaises(ValueError): preflight([self.group(d)],[{'ApplicationShortName':'OTHER'}])
    def test_subnet_overlap(self):
        with self.assertRaises(ValueError): preflight([self.group()],[{'ApplicationShortName':'OTHER','ServerAddresses':'192.0.2.0/24'}])
    def test_batch_overlap(self):
        d=self.base();d['ApplicationName']='SECOND';d['Port']='443'
        with self.assertRaises(ValueError): preflight([self.group(),self.group(d)],[{'ApplicationShortName':'OTHER'}])
    def test_error_body(self):
        with self.assertRaises(ValueError): parse_response(b'<ApplicationConfigurations><NG1APIErrorMessage><ErrorCode>422</ErrorCode></NG1APIErrorMessage></ApplicationConfigurations>')
    def test_verify_settings(self):
        g=self.group();actual=dict(g['fields'],IsActive='true')
        self.assertTrue(verify(g,[actual]))
        actual['ServerPortRange']='443'
        with self.assertRaises(ValueError): verify(g,[actual])
    def test_missing_active(self):
        g=self.group()
        with self.assertRaises(ValueError): verify(g,[g['fields']])
    def test_unknown_parent(self):
        d=self.base();d.update(RegistrationType='CHILD',ParentApplication='HTTPS')
        with self.assertRaises(ValueError): preflight([self.group(d)],[{'ApplicationShortName':'OTHER'}])
    def test_excel_dedup_and_different_ports(self):
        with tempfile.TemporaryDirectory() as t:
            p=Path(t)/'a.xlsx';w=Workbook();s=w.active;s.append(HEADERS)
            d=self.base();s.append([d.get(h,'') for h in HEADERS]);s.append([d.get(h,'') for h in HEADERS]);w.save(p)
            self.assertEqual(len(read_excel(p)),1)
            s.cell(3,HEADERS.index('Port')+1,'443');w.save(p);w.close()
            with self.assertRaises(ValueError):read_excel(p)
    def test_monitor_options(self):
        d=self.base();d.update(RegistrationType='CHILD',ParentApplication='HTTPS',RecordingType='Custom',RecordingSize='128',Fast='1',Expected='2',Degraded='3',ServiceLevel='4',Availability='5')
        self.assertIn('ResponseTimeBoundaries',self.group(d)['fields'])

if __name__=='__main__':unittest.main()
