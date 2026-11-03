"""Explicit application type validation and NETSCOUT XML contracts."""
import ipaddress
import re
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

HEADERS = ['Register', 'ApplicationName', 'RegistrationType', 'ParentApplication', 'Protocol', 'IP', 'Port', 'Description', 'ApplicationGroup', 'ApplicationTag', 'ClientAddresses', 'ClientPortRange', 'URLAddress', 'ExactMatch', 'ErrorCodeFormat', 'Parameter', 'Logged', 'ResponseTime', 'Session', 'RecordingType', 'RecordingSize', 'Fast', 'Expected', 'Degraded', 'ServiceLevel', 'Availability']
INVALID = set('`~!@#$%^&*+=|\\"\'<>?')
MODES = {'SBA': 'Server-based Apps', 'WKA': 'Well Known Apps', 'CHILD': 'Well Known Apps', 'URL': 'URL Application'}
META = {'Application', 'Hostname', 'Found', 'Site', 'Source URL', 'Error', 'Source Sheet', 'Source Row', 'Review'}

def text(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def ports(value):
    spans = []
    for part in value.split(','):
        if not re.fullmatch(r'\d+(?:-\d+)?', part.strip()):
            raise ValueError('Invalid port list/range: ' + value)
        pair = part.strip().split('-'); a, b = int(pair[0]), int(pair[-1])
        if not 1 <= a <= b <= 65535: raise ValueError('Port must be 1-65535')
        spans.append((a, b))
    merged = []
    for a, b in sorted(spans):
        if merged and a <= merged[-1][1] + 1: merged[-1] = (merged[-1][0], max(b, merged[-1][1]))
        else: merged.append((a,b))
    result = ','.join(str(a) if a == b else f'{a}-{b}' for a,b in merged)
    if len(result) > 90: raise ValueError('Port field exceeds 90 characters')
    return result

def networks(value, limit=63, ipv4_only=False, url=False):
    values = set()
    for raw in value.split(','):
        raw = raw.strip()
        n = ipaddress.ip_network(raw, strict=True) if '/' in raw else ipaddress.ip_network(ipaddress.ip_address(raw).exploded)
        if ipv4_only and n.version != 4: raise ValueError('ClientAddresses supports IPv4 only')
        if n.version == 6 and n.prefixlen not in {7,16,32,44,48,64,80,96,*range(112,129)}:
            raise ValueError('Unsupported IPv6 prefix')
        if url and n.version == 4 and n.prefixlen != 32: raise ValueError('URL IPv4 addresses require /32')
        values.add(str(n))
    if len(values) > limit: raise ValueError(f'Address count exceeds {limit}')
    return ','.join(sorted(values))

def row_fields(d):
    mode = d.get('RegistrationType', '').upper()
    if mode not in MODES: raise ValueError('RegistrationType must be SBA, WKA, CHILD or URL')
    name = d.get('ApplicationName','')
    if not name or len(name)>32 or any(c in INVALID or ord(c)<32 for c in name): raise ValueError('Invalid ApplicationName; enter it manually (max 32)')
    parent = d.get('ParentApplication','')
    protocol = d.get('Protocol','').upper()
    if not parent: raise ValueError('ParentApplication must be explicitly entered')
    if mode in ('SBA','WKA'):
        if parent not in ('TCP','UDP','SCTP') or protocol != parent:
            raise ValueError('This prototype requires matching TCP/UDP/SCTP ParentApplication and Protocol')
    if mode == 'CHILD' and parent in ('TCP','UDP','SCTP','IP'):
        raise ValueError('CHILD requires an existing application parent, e.g. HTTP/HTTPS/FTP')
    if mode == 'URL' and parent not in ('HTTP','HTTPS'): raise ValueError('URL parent must be HTTP/HTTPS')
    fields = {'ApplicationShortName': name, 'ApplicationType': MODES[mode], 'ParentApplication': parent}
    allowed = {'Register','ApplicationName','RegistrationType','ParentApplication','Protocol','IP','Port','Description','ApplicationGroup','ApplicationTag'}
    for col, api in [('Description','ApplicationLongName'),('ApplicationGroup','ApplicationGroup'),('ApplicationTag','ApplicationTag')]:
        if d.get(col): fields[api] = d[col]
    if len(d.get('ApplicationTag',''))>1024: raise ValueError('ApplicationTag exceeds 1024 characters')
    if mode == 'SBA':
        allowed |= {'ClientAddresses','ClientPortRange'}
        # Deliberately require address and port to avoid accidental all-host/all-port classification.
        if not d.get('IP') or not d.get('Port'): raise ValueError('SBA requires IP and Port in this tool')
        fields['ServerAddresses'] = networks(d['IP'])
        fields['ServerPortRange'] = ports(d['Port'])
        if d.get('ClientAddresses'): fields['ClientAddresses'] = networks(d['ClientAddresses'],32,True)
        if d.get('ClientPortRange'):
            if not d.get('ClientAddresses'): raise ValueError('ClientPortRange requires ClientAddresses')
            fields['ClientPortRange'] = ports(d['ClientPortRange'])
    elif mode in ('WKA','CHILD'):
        allowed |= {'Logged','ResponseTime','Session','RecordingType','RecordingSize','Fast','Expected','Degraded','ServiceLevel','Availability','Parameter','ErrorCodeFormat'}
        if mode == 'WKA':
            if d.get('IP'): raise ValueError('WKA root does not use IP here; use SBA or CHILD')
            if not re.fullmatch(r'\d+',d.get('Port','')): raise ValueError('WKA requires one ApplicationPort')
            fields['ApplicationPort'] = ports(d['Port'])
        else:
            if not d.get('IP'): raise ValueError('CHILD requires IP in this tool')
            fields['HostAddresses'] = networks(d['IP'], 100000)
            if d.get('Port'): fields['AdditionalPort'] = ports(d['Port'])
        for col, choices in {'Logged':{'Server','Client-Server','None','Reserved'}, 'ResponseTime':{'Enabled','Disabled'}, 'Session':{'Enabled','Disabled'}, 'RecordingType':{'Default','Full','None','Full Optimized','Custom'}, 'ErrorCodeFormat':{'ASCII','Decimal','Default','HEX'}}.items():
            if d.get(col):
                if d[col] not in choices: raise ValueError(f'Invalid {col}')
                fields[col] = d[col]
        if d.get('Parameter'): fields['Parameter'] = d['Parameter']
        if d.get('RecordingType') == 'Custom':
            size = d.get('RecordingSize','')
            if not size.isdigit() or not 0 <= int(size) <= 65535: raise ValueError('Custom recording requires size 0-65535')
            fields['RecordingSize'] = str(int(size))
        elif d.get('RecordingSize'): raise ValueError('RecordingSize requires Custom')
        buckets = ['Fast','Expected','Degraded','ServiceLevel','Availability']
        if any(d.get(k) for k in buckets):
            if not all(d.get(k,'').isdigit() for k in buckets): raise ValueError('Enter all five nonnegative response-time boundaries')
            vals = [int(d[k]) for k in buckets]
            if vals != sorted(vals): raise ValueError('Response-time boundaries must be ascending')
            fields['ResponseTimeBoundaries'] = dict(zip(buckets, map(str, vals)))
    else:
        allowed |= {'URLAddress','ExactMatch','ErrorCodeFormat'}
        if d.get('Description') or d.get('Port'): raise ValueError('URL mode does not support Description or Port here')
        url = d.get('URLAddress','')
        if not url.startswith(parent.lower()+'://'): raise ValueError('URL scheme must match parent')
        fields['URLAddress'] = url
        if d.get('IP'): fields['ServerAddresses'] = networks(d['IP'],48,url=True)
        if d.get('ExactMatch'):
            if d['ExactMatch'].lower() not in ('true','false'): raise ValueError('ExactMatch must be true/false')
            fields['ExactMatch'] = d['ExactMatch'].lower()
        if d.get('ErrorCodeFormat'):
            if d['ErrorCodeFormat'] not in ('ASCII','Decimal','Default','HEX'): raise ValueError('Invalid ErrorCodeFormat')
            fields['ErrorCodeFormat'] = d['ErrorCodeFormat']
    for key,value in d.items():
        if value and key not in allowed|META: raise ValueError(f'Unsupported option for {mode}: {key}')
    return mode, fields

def read_excel(path):
    w=load_workbook(path, read_only=True, data_only=False)
    try:
        s=w['Registration'] if 'Registration' in w.sheetnames else w.active
        iterator=s.iter_rows(); headers=[text(c.value) for c in next(iterator)]
        if len(set(headers))!=len(headers): raise ValueError('Duplicate headers')
        if not {'Register','ApplicationName','RegistrationType','ParentApplication'} <= set(headers): raise ValueError('Use the new registration template with RegistrationType and ParentApplication')
        groups={}
        for number,cells in enumerate(iterator,2):
            d=dict(zip(headers,[text(c.value) for c in cells])); flag=d.get('Register','').upper()
            if flag in ('','N'): continue
            if flag!='Y': raise ValueError(f'Row {number}: Register must be Y/N')
            if any(c.data_type=='f' for c in cells): raise ValueError(f'Row {number}: formulas not allowed')
            try: mode,fields=row_fields(d)
            except ValueError as exc: raise ValueError(f'Row {number}: {exc}') from exc
            key=fields['ApplicationShortName'].casefold()
            if key in groups:
                # Merge only equivalent address-independent configurations, never expand port cross-products.
                g=groups[key]; address='HostAddresses' if mode=='CHILD' else 'ServerAddresses'
                a={k:v for k,v in fields.items() if k!=address}
                b={k:v for k,v in g['fields'].items() if k!=address}
                if mode!=g['mode'] or a!=b: raise ValueError(f'Row {number}: same name has different settings/ports')
                if address in fields:
                    limit=48 if mode=='URL' else 63 if mode=='SBA' else 100000
                    g['fields'][address]=networks(g['fields'][address]+','+fields[address],limit,url=mode=='URL')
                g['rows'].append(number)
            else: groups[key]={'name':fields['ApplicationShortName'],'mode':mode,'fields':fields,'rows':[number]}
        return list(groups.values())
    finally: w.close()

def xml_for(g):
    root=ET.Element('ApplicationConfigurations'); app=ET.SubElement(root,'ApplicationConfiguration')
    for key,value in g['fields'].items():
        el=ET.SubElement(app,key)
        if isinstance(value,dict):
            for k,v in value.items(): ET.SubElement(el,k).text=v
        else: el.text=value
    ET.indent(root)
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)

def parse_response(data):
    root=ET.fromstring(data)
    for el in root.iter(): el.tag=el.tag.split('}')[-1]
    errors=[el for el in root.iter() if el.tag in ('NG1APIErrorMessage','ErrorCode','errorMessageList','errorCode')]
    if errors: raise ValueError('NETSCOUT error response: '+ ' '.join(root.itertext())[:1600])
    if root.tag!='ApplicationConfigurations': raise ValueError('Unexpected XML response root')
    apps=[]
    for app in root.findall('ApplicationConfiguration'):
        fields={el.tag: ({c.tag:text(c.text) for c in el} if len(el) else text(el.text)) for el in app}
        if not fields.get('ApplicationShortName'): raise ValueError('Application missing name')
        apps.append(fields)
    return apps

def preflight(groups, existing):
    if not existing: raise ValueError('Inventory empty; stop for review')
    seen=list(existing)
    names={a['ApplicationShortName'].casefold() for a in existing}
    for g in groups:
        f=g['fields']; name=g['name']
        if name.casefold() in names: raise ValueError('Existing name: '+name)
        # Parent must already exist: pending-parent creation order is deliberately unsupported.
        parent=f['ParentApplication']
        if parent not in ('TCP','UDP','SCTP','IP') and parent.casefold() not in names:
            raise ValueError('Parent missing from inventory: '+parent)
        if g['mode'] in ('SBA','URL') and f.get('ServerAddresses'):
            for a in seen:
                if not a.get('ServerAddresses'): continue
                for left in f['ServerAddresses'].split(','):
                    n=ipaddress.ip_network(left)
                    for right in a['ServerAddresses'].split(','):
                        other=ipaddress.ip_network(right.strip(),strict=False)
                        if n.version==other.version and n.overlaps(other): raise ValueError(f"Server address overlap: {name} / {a['ApplicationShortName']}")
        if f.get('ClientAddresses'):
            for left in f.get('ServerAddresses','').split(','):
                if not left: continue
                n=ipaddress.ip_network(left)
                for right in f['ClientAddresses'].split(','):
                    other=ipaddress.ip_network(right)
                    if n.version==other.version and n.overlaps(other): raise ValueError('Server/client address overlap: '+name)
        if g['mode']=='URL' and any(a.get('URLAddress')==f['URLAddress'] for a in seen): raise ValueError('Existing URL: '+f['URLAddress'])
        seen.append(f)

def verify(g, apps):
    matches=[a for a in apps if a['ApplicationShortName'].casefold()==g['name'].casefold()]
    if len(matches)!=1: raise ValueError('Expected exactly one application after registration')
    actual=matches[0]; differences=[]
    for key,wanted in g['fields'].items():
        got=actual.get(key)
        if key in ('ServerAddresses','HostAddresses','ClientAddresses'):
            try: same=networks(wanted,100000)==networks(got or '',100000)
            except ValueError: same=False
        elif key in ('ServerPortRange','ClientPortRange','AdditionalPort','ApplicationPort'):
            try: same=ports(wanted)==ports(got or '')
            except ValueError: same=False
        else: same=got==wanted
        if not same: differences.append(key)
    if text(actual.get('IsActive')).lower()!='true': differences.append('IsActive (missing/false)')
    if differences: raise ValueError('Verification requires review: '+', '.join(differences))
    return True
