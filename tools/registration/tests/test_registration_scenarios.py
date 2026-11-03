"""Ten representative NETSCOUT synchronization scenarios."""
import unittest

from registration_rules import row_fields
from sync_planner import plan_sync


def desired(name='APP_A', mode='SBA', parent='TCP', protocol='TCP',
            ip='192.0.2.10', port='8443'):
    row = dict(Register='Y', ApplicationName=name, RegistrationType=mode,
               ParentApplication=parent, Protocol=protocol, IP=ip, Port=port)
    resolved_mode, fields = row_fields(row)
    return {'name': name, 'mode': resolved_mode, 'fields': fields, 'rows': [2]}


def actual(name='APP_A', app_type='Server-based Apps', parent='TCP',
           ip='192.0.2.10/32', port='8443'):
    return {'ApplicationShortName': name, 'ApplicationType': app_type,
            'ParentApplication': parent, 'ServerAddresses': ip,
            'ServerPortRange': port, 'IsActive': 'true'}


class TenSyncScenarios(unittest.TestCase):
    def assert_action(self, expected, group, inventory):
        item = plan_sync([group], inventory)[0]
        self.assertEqual(item['action'], expected, item)
        return item

    def test_01_new_application_is_create(self):
        self.assert_action('CREATE', desired(), [{'ApplicationShortName': 'HTTP'}])

    def test_02_identical_application_is_matched(self):
        self.assert_action('MATCHED', desired(), [actual()])

    def test_03_new_ip_is_add_ip(self):
        item = self.assert_action('ADD_IP', desired(ip='192.0.2.10,192.0.2.11'), [actual()])
        self.assertEqual(item['update_group']['fields']['ServerAddresses'],
                         '192.0.2.10/32,192.0.2.11/32')
        self.assertEqual(item['update_group']['fields']['ApplicationType'], 'Server-based Apps')
        self.assertEqual(item['update_group']['fields']['ParentApplication'], 'TCP')
        self.assertEqual(item['update_group']['fields']['ServerPortRange'], '8443')

    def test_04_new_port_is_add_port(self):
        item = self.assert_action('ADD_PORT', desired(port='8443,9443'), [actual()])
        self.assertEqual(item['update_group']['fields']['ServerPortRange'], '8443,9443')

    def test_05_new_ip_and_port_are_combined(self):
        item = self.assert_action('ADD_IP+ADD_PORT',
                                  desired(ip='192.0.2.10,192.0.2.11', port='8443,9443'),
                                  [actual()])
        self.assertIn('192.0.2.11/32', item['update_group']['fields']['ServerAddresses'])
        self.assertEqual(item['update_group']['fields']['ServerPortRange'], '8443,9443')

    def test_06_address_owned_by_other_app_is_conflict(self):
        self.assert_action('CONFLICT', desired(), [actual(name='OTHER_APP')])

    def test_07_parent_mismatch_is_conflict(self):
        self.assert_action('CONFLICT', desired(), [actual(parent='UDP')])

    def test_08_application_type_mismatch_is_conflict(self):
        self.assert_action('CONFLICT', desired(), [actual(app_type='Well Known Apps')])

    def test_09_identical_wka_is_matched(self):
        group = desired(name='WKA_A', mode='WKA', ip='', port='55004')
        inventory = [{'ApplicationShortName': 'WKA_A', 'ApplicationType': 'Well Known Apps',
                      'ParentApplication': 'TCP', 'ApplicationPort': '55004', 'IsActive': 'true'}]
        self.assert_action('MATCHED', group, inventory)

    def test_10_wka_port_change_is_conflict(self):
        group = desired(name='WKA_A', mode='WKA', ip='', port='55005')
        inventory = [{'ApplicationShortName': 'WKA_A', 'ApplicationType': 'Well Known Apps',
                      'ParentApplication': 'TCP', 'ApplicationPort': '55004', 'IsActive': 'true'}]
        self.assert_action('CONFLICT', group, inventory)


if __name__ == '__main__':
    unittest.main(verbosity=2)
