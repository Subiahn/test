"""Plan and optionally apply safe NETSCOUT application creates and additive updates."""
import argparse
import getpass
import json
import os
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
for dependency_path in (PROJECT_ROOT / 'vendor', PROJECT_ROOT / 'src'):
    if str(dependency_path) not in sys.path:
        sys.path.insert(0, str(dependency_path))

from netscout_api import NetscoutClient
from registration_rules import parse_response, read_excel, verify, xml_for
from sync_planner import plan_sync


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True)
    parser.add_argument('--base-url', required=True)
    parser.add_argument('--username', required=True)
    parser.add_argument('--apply', action='store_true', help='Apply CREATE and additive PUT actions')
    parser.add_argument('--ca-file')
    parser.add_argument('--insecure', action='store_true')
    args = parser.parse_args()
    password = os.environ.get('NETSCOUT_PASSWORD') or getpass.getpass('NETSCOUT password: ')
    out = Path('results') / ('sync_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f'))
    out.mkdir(parents=True)
    def http_error(status, body):
        (out / 'last_http_error.txt').write_bytes(body)
    client = NetscoutClient(args.base_url, args.username, password, args.ca_file, args.insecure, http_error)
    groups = read_excel(args.input)
    inventory = parse_response(client.get_all())
    plan = plan_sync(groups, inventory)
    public = [{k: item[k] for k in ('action', 'name', 'message')} for item in plan]
    (out / 'plan.json').write_text(json.dumps(public, ensure_ascii=False, indent=2), encoding='utf-8')
    for item in public:
        print(f"{item['action']:<16} {item['name']}: {item['message']}")
    actionable = [x for x in plan if x['action'] == 'CREATE' or x['action'].startswith('ADD_')]
    if not args.apply:
        print(f'Plan only; {len(actionable)} applicable change(s). Results: {out}')
        return
    conflicts = [x for x in plan if x['action'] == 'CONFLICT']
    if conflicts:
        raise ValueError('Conflicts exist; no changes applied')
    if not actionable:
        print('No changes required.')
        return
    if input('Type SYNC to apply the plan: ') != 'SYNC':
        print('Cancelled; no writes performed.')
        return
    results = []
    for item in actionable:
        record = {'action': item['action'], 'name': item['name'], 'status': 'UNKNOWN'}
        results.append(record)
        (out / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
        payload_group = item['group'] if item['action'] == 'CREATE' else item['update_group']
        payload = xml_for(payload_group)
        (out / f"{len(results):04d}_{item['action']}.xml").write_bytes(payload)
        if item['action'] == 'CREATE':
            client.create(payload)
        else:
            client.update(item['name'], payload)
        current = parse_response(client.get_all())
        verify(item.get('verify_group', item['group']), current)
        record['status'] = 'VERIFIED'
        (out / 'results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'Applied and verified {len(results)} change(s). Results: {out}')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'ERROR: {exc}')
        raise SystemExit(1)
