"""Convert hostname lookup Excel to registration Excel; never calls an API.
Python 3.10+, openpyxl. No application names are generated.
"""
import sys
from pathlib import Path
PROJECT_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(PROJECT_ROOT / 'vendor'))
import argparse
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from registration_rules import HEADERS, text

EXTRA = ['Application', 'Hostname', 'Found', 'Site', 'Source URL', 'Error', 'Source Sheet', 'Source Row', 'Review']
ALIASES = {'ip': 'IP', 'ipaddress': 'IP', 'port': 'Port', 'serverport': 'Port',
           'application': 'Application', 'app': 'Application', 'applications': 'Application',
           'hostname': 'Hostname', 'found': 'Found', 'site': 'Site',
           'sourceurl': 'Source URL', 'error': 'Error', 'protocol': 'Protocol', 'transport': 'Protocol'}

def header(v):
    raw = re.sub(r'[\s_\-]+', '', text(v).lower())
    return ALIASES.get(raw, text(v))

def convert(source, output, sheet=None, protocol=None):
    source, output = Path(source), Path(output)
    if source.resolve() == output.resolve():
        raise ValueError('Input and output must be different files')
    if output.exists():
        raise ValueError('Output already exists; use a new path to preserve manual edits')
    wb = load_workbook(source, read_only=True, data_only=False)
    try:
        eligible = []
        for s in wb.worksheets:
            headers = [header(c.value) for c in next(s.iter_rows(), [])]
            if {'IP', 'Port', 'Hostname'} <= set(headers):
                eligible.append(s.title)
        if sheet:
            if sheet not in eligible:
                raise ValueError('Selected sheet must contain IP, Port and Hostname headers in row 1')
        elif len(eligible) == 1:
            sheet = eligible[0]
        else:
            raise ValueError('Specify --sheet; eligible sheets: ' + ', '.join(eligible))
        ws = wb[sheet]
        iterator = ws.iter_rows()
        headers = [header(c.value) for c in next(iterator)]
        if len(set(headers)) != len(headers):
            raise ValueError('Duplicate or aliased duplicate headers')
        result = Workbook()
        dest = result.active
        dest.title = 'Registration'
        dest.append(HEADERS + EXTRA)
        skipped, count = 0, 0
        for row_number, cells in enumerate(iterator, 2):
            d = dict(zip(headers, [text(c.value) for c in cells]))
            host = d.get('Hostname', '')
            if not host or host.casefold() in ('none', 'null', 'nan', 'n/a', '-', 'not found'):
                skipped += 1
                continue
            # Formula-backed lookup outputs must be exported as values first.
            if any(c.data_type == 'f' for c in cells):
                raise ValueError(f'Row {row_number}: formulas found; provide a values-only lookup export')
            transport = d.get('Protocol', '').upper()
            review = []
            if protocol:
                transport = protocol
            if transport not in ('TCP', 'UDP'):
                transport = ''
                review.append('Protocol 직접 입력 필요')
            found = d.get('Found', '')
            if found.casefold() in ('false', 'n', 'no', '0'):
                review.append('Hostname이 있지만 Found가 실패 상태')
            if d.get('Error'):
                review.append('조회 Error 확인 필요')
            if not d.get('IP') or not d.get('Port'):
                review.append('IP 또는 Port 누락')
            # Keep every source row/site result for review; registration reader deduplicates rules.
            values = {'Register':'N', 'Protocol':transport, 'IP':d.get('IP',''), 'Port':d.get('Port',''),
                      'Application':d.get('Application',''), 'Hostname':host, 'Found':found, 'Site':d.get('Site',''),
                      'Source URL':d.get('Source URL',''), 'Error':d.get('Error',''), 'Source Sheet':sheet,
                      'Source Row':row_number, 'Review':'; '.join(review + ['RegistrationType 및 ParentApplication 직접 입력 필요'])}
            dest.append([values.get(k,'') for k in HEADERS+EXTRA])
            # Treat external content as literal text, never as Excel formulas.
            for cell in dest[dest.max_row]:
                if isinstance(cell.value, str):
                    cell.data_type = 's'
            count += 1
        dest.freeze_panes = 'A2'
        dest.auto_filter.ref = dest.dimensions
        for cell in dest[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='16324F')
        for column in dest.columns:
            dest.column_dimensions[column[0].column_letter].width = 23
        dest.column_dimensions[dest.cell(1,len(HEADERS+EXTRA)).column_letter].width = 55
        for col, formula in [('A', '"Y,N"'), ('C', '"SBA,WKA,CHILD,URL"'), ('E', '"TCP,UDP,SCTP"')]:
            validation = DataValidation(type='list', formula1=formula, allow_blank=True)
            validation.errorTitle = '허용값 확인'
            validation.error = '목록에서 값을 선택하세요.'
            validation.showErrorMessage = True
            validation.errorStyle = 'stop'
            dest.add_data_validation(validation)
            validation.add(f'{col}2:{col}{max(2, dest.max_row)}')
        for row in dest.iter_rows(min_row=2):
            for index in (0, 1, 2, 3, 4, 7):
                row[index].fill = PatternFill('solid', fgColor='FFF2CC')
        output.parent.mkdir(parents=True, exist_ok=True)
        result.save(output)
        result.close()
        return {'written': count, 'skipped_no_hostname': skipped, 'sheet': sheet}
    finally:
        wb.close()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True, help='Hostname lookup result .xlsx')
    parser.add_argument('--output', default='registration.xlsx')
    parser.add_argument('--sheet', help='Required when more than one matching sheet exists')
    parser.add_argument('--protocol', choices=['TCP', 'UDP'], help='Set only if ALL selected source records use this transport')
    args = parser.parse_args()
    print(convert(args.input, args.output, args.sheet, args.protocol))

if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f'ERROR: {exc}', file=sys.stderr)
        raise SystemExit(1)
