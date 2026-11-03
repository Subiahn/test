import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook

import netscout_sync
from registration_rules import HEADERS, row_fields
from sync_planner import plan_sync


class Response:
    def __init__(self, body, status=200):
        self.body, self.status = body, status
    def __enter__(self): return self
    def __exit__(self, *args): pass
    def read(self): return self.body


class SyncPlannerTests(unittest.TestCase):
    def group(self, ip='192.0.2.10', port='8443'):
        row = dict(Register='Y', ApplicationName='APP_A', RegistrationType='SBA',
                   ParentApplication='TCP', Protocol='TCP', IP=ip, Port=port)
        mode, fields = row_fields(row)
        return {'name': 'APP_A', 'mode': mode, 'fields': fields, 'rows': [2]}

    def existing(self, ip='192.0.2.10/32', port='8443'):
        return {'ApplicationShortName': 'APP_A', 'ApplicationType': 'Server-based Apps',
                'ParentApplication': 'TCP', 'ServerAddresses': ip,
                'ServerPortRange': port, 'IsActive': 'true'}

    def test_create(self):
        self.assertEqual(plan_sync([self.group()], [{'ApplicationShortName': 'HTTP'}])[0]['action'], 'CREATE')

    def test_matched(self):
        self.assertEqual(plan_sync([self.group()], [self.existing()])[0]['action'], 'MATCHED')

    def test_add_ip_preserves_existing(self):
        item = plan_sync([self.group('192.0.2.10,192.0.2.11')], [self.existing()])[0]
        self.assertEqual(item['action'], 'ADD_IP')
        self.assertEqual(item['update_group']['fields']['ServerAddresses'], '192.0.2.10/32,192.0.2.11/32')
        self.assertEqual(item['update_group']['fields']['ApplicationType'], 'Server-based Apps')
        self.assertEqual(item['update_group']['fields']['ParentApplication'], 'TCP')
        self.assertEqual(item['update_group']['fields']['ServerPortRange'], '8443')

    def test_add_port_preserves_existing(self):
        item = plan_sync([self.group(port='8443,9443')], [self.existing()])[0]
        self.assertEqual(item['action'], 'ADD_PORT')
        self.assertEqual(item['update_group']['fields']['ServerPortRange'], '8443,9443')

    def test_other_app_address_conflict(self):
        other = self.existing(); other['ApplicationShortName'] = 'OTHER'
        self.assertEqual(plan_sync([self.group()], [other])[0]['action'], 'CONFLICT')

    def test_apply_add_ip_uses_put_and_verifies(self):
        with tempfile.TemporaryDirectory() as temp:
            workbook = Path(temp) / 'sync.xlsx'
            w = Workbook(); sheet = w.active; sheet.append(HEADERS)
            row = dict(Register='Y', ApplicationName='APP_A', RegistrationType='SBA',
                       ParentApplication='TCP', Protocol='TCP',
                       IP='192.0.2.10,192.0.2.11', Port='8443')
            sheet.append([row.get(h, '') for h in HEADERS]); w.save(workbook); w.close()
            before = b'''<ApplicationConfigurations><ApplicationConfiguration>
                <ApplicationShortName>APP_A</ApplicationShortName><ApplicationType>Server-based Apps</ApplicationType>
                <ParentApplication>TCP</ParentApplication><ServerAddresses>192.0.2.10/32</ServerAddresses>
                <ServerPortRange>8443</ServerPortRange><IsActive>true</IsActive>
                </ApplicationConfiguration></ApplicationConfigurations>'''
            after = before.replace(b'192.0.2.10/32', b'192.0.2.10/32,192.0.2.11/32')
            opener = Mock(); opener.open.side_effect = [Response(before), Response(after), Response(after)]
            old = os.getcwd()
            try:
                os.chdir(temp)
                argv = ['netscout_sync.py', '--input', str(workbook), '--base-url',
                        'https://example.test:8443', '--username', 'tester', '--apply']
                with patch.object(sys, 'argv', argv), patch.dict(os.environ, {'NETSCOUT_PASSWORD': 'test-only'}), \
                     patch('builtins.input', return_value='SYNC'), \
                     patch('urllib.request.build_opener', return_value=opener):
                    netscout_sync.main()
                requests = [call.args[0] for call in opener.open.call_args_list]
                self.assertEqual([r.get_method() for r in requests], ['GET', 'PUT', 'GET'])
                self.assertTrue(requests[1].full_url.endswith('/ng1api/ncm/applications/APP_A'))
                self.assertIn(b'192.0.2.10/32,192.0.2.11/32', requests[1].data)
                self.assertIn(b'<ApplicationType>Server-based Apps</ApplicationType>', requests[1].data)
                self.assertIn(b'<ParentApplication>TCP</ParentApplication>', requests[1].data)
                self.assertIn(b'<ServerPortRange>8443</ServerPortRange>', requests[1].data)
            finally:
                os.chdir(old)


if __name__ == '__main__':
    unittest.main()
