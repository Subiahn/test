"""IP/Port 후보 YAML을 기반으로 내부 Hostname 조회 및 Excel 보고서를 생성한다.

사용 예:
  python tools/lookup_hostnames.py --input output/ip_port_candidates.yaml
  python tools/lookup_hostnames.py --input output/ip_port_candidates.yaml --config config/hostname_lookup.yaml

설정 YAML 예:
  timeout: 10
  search_param_name: search
  verify_ssl: false
  sites:
    - name: Site1
      url: http://SITE1_SERVER:5000/Search
      type: table
      hostname_index: 0
      ip_index: 7
    - name: Site2
      url: http://SITE2_SERVER:5002/Search
      type: input
"""
from __future__ import annotations

import argparse
import ipaddress
import sys
from collections import Counter
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
VENDOR = ROOT / 'vendor'

try:
    import requests
    import yaml
    import urllib3
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ModuleNotFoundError:
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    import requests
    import yaml
    import urllib3
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill


class TableParser(HTMLParser):
    """테스트에서 검증된 방식: 모든 tr/td 셀을 순서대로 수집."""

    def __init__(self):
        super().__init__()
        self.in_tr = False
        self.in_td = False
        self.current_row: list[str] = []
        self.current_text = ''
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        if tag == 'tr':
            self.in_tr = True
            self.in_td = False
            self.current_row = []
        elif tag == 'td' and self.in_tr:
            self.in_td = True
            self.current_text = ''

    def handle_data(self, data):
        if self.in_td:
            self.current_text += data

    def handle_endtag(self, tag):
        tag = tag.lower()
        if tag == 'td' and self.in_td:
            self.current_row.append(self.current_text.strip())
            self.current_text = ''
            self.in_td = False
        elif tag == 'tr' and self.in_tr:
            if self.current_row:
                self.rows.append(self.current_row)
            self.current_row = []
            self.in_tr = False
            self.in_td = False


class InputParser(HTMLParser):
    """테스트에서 검증된 방식: 같은 tr 안의 input name=ip/hostname 연결."""

    def __init__(self):
        super().__init__()
        self.in_tr = False
        self.current_ip = ''
        self.current_hostname = ''
        self.results: list[dict[str, str]] = []

    def handle_starttag(self, tag, attrs):
        tag = tag.lower()
        attrs_dict = {str(k).lower(): str(v or '') for k, v in attrs}
        if tag == 'tr':
            self.in_tr = True
            self.current_ip = ''
            self.current_hostname = ''
        elif tag == 'input' and self.in_tr:
            name = attrs_dict.get('name', '').strip().lower()
            value = attrs_dict.get('value', '').strip()
            if name == 'ip':
                self.current_ip = value
            elif name == 'hostname':
                self.current_hostname = value

    def handle_endtag(self, tag):
        if tag.lower() == 'tr' and self.in_tr:
            if self.current_ip:
                self.results.append({'ip': self.current_ip, 'hostname': self.current_hostname})
            self.in_tr = False


def as_applications(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (list, tuple, set)):
        return list(dict.fromkeys(str(v).strip() for v in value if str(v).strip()))
    if isinstance(value, str):
        return list(dict.fromkeys(v.strip() for v in value.split(',') if v.strip()))
    return [str(value).strip()]


def normalize_candidate(item: Any) -> dict[str, str] | None:
    if isinstance(item, str):
        item = {'ip': item}
    if not isinstance(item, dict):
        return None
    lower = {str(k).lower(): v for k, v in item.items()}
    ip = str(lower.get('ip') or lower.get('ip_address') or '').strip().lstrip('+').strip()
    try:
        ipaddress.ip_address(ip)
    except ValueError:
        return None
    port = lower.get('port', '')
    port = '' if port is None else str(port).strip()
    apps = as_applications(lower.get('applications', lower.get('application', lower.get('app'))))
    return {'IP': ip, 'Port': port, 'Application': ', '.join(apps)}


def load_candidates(path: Path) -> list[dict[str, str]]:
    with path.open('r', encoding='utf-8-sig') as f:
        data = yaml.safe_load(f)
    if isinstance(data, dict):
        raw = data.get('unique_ip_ports')
        if not raw:
            raw = data.get('unique_ips', [])
    elif isinstance(data, list):
        raw = data
    else:
        raw = []
    if not isinstance(raw, list):
        raise ValueError('YAML unique_ip_ports/unique_ips는 리스트여야 합니다.')
    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()
    for item in raw:
        row = normalize_candidate(item)
        if not row:
            print(f'[WARN] 잘못된 IP 또는 행 제외: {item!r}')
            continue
        key = (row['IP'], row['Port'], row['Application'])
        if key not in seen:
            seen.add(key)
            rows.append(row)
    return rows


def load_config(path: Path) -> dict[str, Any]:
    with path.open('r', encoding='utf-8-sig') as f:
        config = yaml.safe_load(f)
    if not isinstance(config, dict):
        raise ValueError('설정 YAML은 매핑 형태여야 합니다.')
    sites = config.get('sites')
    if not isinstance(sites, list) or not sites:
        raise ValueError('설정 YAML에 sites 리스트가 필요합니다.')
    for site in sites:
        if not isinstance(site, dict) or not site.get('url'):
            raise ValueError('각 사이트에 url이 필요합니다.')
        if str(site.get('type', '')).lower() not in ('table', 'input'):
            raise ValueError('사이트 type은 table 또는 input이어야 합니다.')
    return config


def lookup_one_ip(ip: str, config: dict[str, Any]) -> dict[str, Any]:
    errors: list[str] = []
    for site in config['sites']:
        name = str(site.get('name', 'Site'))
        url = str(site['url'])
        site_type = str(site['type']).lower()
        search_param = str(site.get('search_param_name', config.get('search_param_name', 'search')))
        timeout = float(site.get('timeout', config.get('timeout', 10)))
        verify_ssl = bool(site.get('verify_ssl', config.get('verify_ssl', True)))
        try:
            if not verify_ssl:
                urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            response = requests.get(url, params={search_param: ip}, timeout=timeout, verify=verify_ssl)
            print(f'[REQUEST] IP={ip} SITE={name} URL={response.url} HTTP={response.status_code}')
            response.raise_for_status()
            found_hostname = ''
            matched_ip = False
            if site_type == 'table':
                parser = TableParser()
                parser.feed(response.text)
                parser.close()
                hostname_index = int(site.get('hostname_index', 0))
                ip_index = int(site.get('ip_index', 7))
                if hostname_index < 0 or ip_index < 0:
                    raise ValueError('인덱스는 0 이상이어야 합니다.')
                for row in parser.rows:
                    if len(row) <= max(hostname_index, ip_index):
                        continue
                    if row[ip_index].strip() == ip:
                        matched_ip = True
                        found_hostname = row[hostname_index].strip()
                        if found_hostname:
                            break
                print(f'[PARSE] SITE={name} rows={len(parser.rows)} ip_match={matched_ip} hostname={found_hostname!r}')
            else:
                parser = InputParser()
                parser.feed(response.text)
                parser.close()
                for row in parser.results:
                    if row['ip'].strip() == ip:
                        matched_ip = True
                        found_hostname = row['hostname'].strip()
                        if found_hostname:
                            break
                print(f'[PARSE] SITE={name} rows={len(parser.results)} ip_match={matched_ip} hostname={found_hostname!r}')
            if found_hostname:
                return {'Found': True, 'Hostname': found_hostname, 'Site': name,
                        'Source URL': response.url, 'Error': ''}
            errors.append(f'{name}: ' + ('Hostname empty' if matched_ip else 'IP not found'))
        except (requests.exceptions.RequestException, ValueError) as exc:
            errors.append(f'{name}: {exc}')
            print(f'[ERROR] IP={ip} SITE={name} {exc}')
    return {'Found': False, 'Hostname': '', 'Site': '', 'Source URL': '', 'Error': ' | '.join(errors)}


def write_excel(candidates: list[dict[str, str]], lookup_cache: dict[str, dict[str, Any]], output: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Hostname_Lookup'
    headers = ['IP', 'Port', 'Application', 'Found', 'Hostname', 'Site', 'Source URL', 'Error']
    ws.append(headers)
    for candidate in candidates:
        lookup = lookup_cache[candidate['IP']]
        ws.append([candidate['IP'], candidate['Port'], candidate['Application'],
                   lookup['Found'], lookup['Hostname'], lookup['Site'],
                   lookup['Source URL'], lookup['Error']])
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    summary = wb.create_sheet('Run_Summary')
    summary.append(['Metric', 'Value'])
    found = sum(bool(result['Found']) for result in lookup_cache.values())
    counts = Counter(candidate['IP'] for candidate in candidates)
    metrics = [
        ('run_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        ('ip_count', len(lookup_cache)),
        ('found_count', found),
        ('not_found_count', len(lookup_cache) - found),
        ('result_row_count', len(candidates)),
        ('unique_ip_port_count', len({(row['IP'], row['Port']) for row in candidates})),
        ('duplicate_ip_count', sum(1 for count in counts.values() if count > 1)),
    ]
    for item in metrics:
        summary.append(item)
    for sheet in (ws, summary):
        for cell in sheet[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='25466B')
            cell.alignment = Alignment(horizontal='center')
        for column, width in {'A': 20, 'B': 14, 'C': 30, 'D': 12, 'E': 36,
                              'F': 18, 'G': 68, 'H': 55}.items():
            sheet.column_dimensions[column].width = width
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)


def default_output_path(input_path: Path) -> Path:
    """v0.86 후보 폴더 구조면 05_hostname/<app>/hostname_lookup.xlsx에 저장한다."""
    parent = input_path.parent
    if parent.parent.name == "04_candidates":
        return parent.parent.parent / "05_hostname" / parent.name / "hostname_lookup.xlsx"
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    return parent / f'hostname_lookup_{timestamp}.xlsx'


def main() -> int:
    parser = argparse.ArgumentParser(description='YAML IP/Port 후보 → Hostname 조회 → Excel')
    parser.add_argument('--input', default='output/ip_port_candidates.yaml', help='후보 YAML 경로')
    parser.add_argument('--config', default='config/hostname_lookup.yaml', help='사이트 설정 YAML 경로')
    parser.add_argument('--output', help='결과 XLSX 경로 (기본: v0.86 구조면 05_hostname/<app>/hostname_lookup.xlsx)')
    args = parser.parse_args()
    input_path = Path(args.input)
    config_path = Path(args.config)
    if not input_path.is_file():
        print(f'[ERROR] 입력 파일 없음: {input_path}')
        return 1
    if not config_path.is_file():
        print(f'[ERROR] 설정 파일 없음: {config_path}')
        return 1
    try:
        candidates = load_candidates(input_path)
        config = load_config(config_path)
        if not candidates:
            print('[ERROR] 유효한 IP 후보가 없습니다.')
            return 1
        unique_ips = list(dict.fromkeys(row['IP'] for row in candidates))
        print(f'[START] candidates={len(candidates)} unique_ips={len(unique_ips)}')
        lookup_cache = {}
        for i, ip in enumerate(unique_ips, 1):
            print(f'[{i}/{len(unique_ips)}] {ip}')
            lookup_cache[ip] = lookup_one_ip(ip, config)
        output = Path(args.output) if args.output else default_output_path(input_path)
        write_excel(candidates, lookup_cache, output)
        print(f'[DONE] Excel={output} found={sum(v["Found"] for v in lookup_cache.values())}/{len(unique_ips)}')
        return 0
    except (OSError, ValueError, yaml.YAMLError) as exc:
        print(f'[ERROR] {exc}')
        return 1


if __name__ == '__main__':
    sys.exit(main())
