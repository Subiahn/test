"""Pure desired-state comparison for NETSCOUT application synchronization."""
import ipaddress

from registration_rules import networks, ports

ADDRESS_FIELD = {'SBA': 'ServerAddresses', 'CHILD': 'HostAddresses', 'URL': 'ServerAddresses'}
PORT_FIELD = {'SBA': 'ServerPortRange', 'CHILD': 'AdditionalPort', 'WKA': 'ApplicationPort'}


def _network_set(value):
    if not value:
        return set()
    return {str(ipaddress.ip_network(x.strip(), strict=False)) for x in value.split(',') if x.strip()}


def _port_set(value):
    result = set()
    if not value:
        return result
    for part in value.split(','):
        pair = part.strip().split('-')
        result.update(range(int(pair[0]), int(pair[-1]) + 1))
    return result


def _format_ports(values):
    if not values:
        return ''
    spans = []
    for value in sorted(values):
        if spans and value == spans[-1][1] + 1:
            spans[-1] = (spans[-1][0], value)
        else:
            spans.append((value, value))
    return ports(','.join(str(a) if a == b else f'{a}-{b}' for a, b in spans))


def _overlap_owner(addresses, existing, own_name=None):
    for app in existing:
        if own_name and app.get('ApplicationShortName', '').casefold() == own_name.casefold():
            continue
        for key in ('ServerAddresses', 'HostAddresses'):
            for right in _network_set(app.get(key, '')):
                other = ipaddress.ip_network(right)
                for left in addresses:
                    candidate = ipaddress.ip_network(left)
                    if candidate.version == other.version and candidate.overlaps(other):
                        return app.get('ApplicationShortName', '?')
    return None


def plan_sync(groups, existing):
    """Return a side-effect-free plan. Updates only add addresses/ports."""
    by_name = {a['ApplicationShortName'].casefold(): a for a in existing}
    plan = []
    for group in groups:
        name, mode, desired = group['name'], group['mode'], group['fields']
        current = by_name.get(name.casefold())
        address_key, port_key = ADDRESS_FIELD.get(mode), PORT_FIELD.get(mode)
        desired_addresses = _network_set(desired.get(address_key, '')) if address_key else set()
        owner = _overlap_owner(desired_addresses, existing, name if current else None)
        if owner:
            plan.append({'action': 'CONFLICT', 'name': name, 'group': group,
                         'message': f'Address overlaps application {owner}'})
            continue
        if current is None:
            plan.append({'action': 'CREATE', 'name': name, 'group': group, 'message': 'Application does not exist'})
            continue
        incompatible = []
        for key in ('ApplicationType', 'ParentApplication'):
            if desired.get(key) and current.get(key) != desired[key]:
                incompatible.append(key)
        for key, value in desired.items():
            if key in {address_key, port_key, 'ApplicationShortName', 'ApplicationType', 'ParentApplication'}:
                continue
            if value and current.get(key) != value:
                incompatible.append(key)
        if incompatible:
            plan.append({'action': 'CONFLICT', 'name': name, 'group': group,
                         'message': 'Different existing settings: ' + ', '.join(sorted(set(incompatible)))})
            continue
        changes = {}
        labels = []
        if address_key:
            current_addresses = _network_set(current.get(address_key, ''))
            added = desired_addresses - current_addresses
            if added:
                limit = 48 if mode == 'URL' else 63 if mode == 'SBA' else 100000
                changes[address_key] = networks(','.join(sorted(current_addresses | desired_addresses)), limit, url=mode == 'URL')
                labels.append('ADD_IP')
        if port_key:
            wanted_ports = _port_set(desired.get(port_key, ''))
            current_ports = _port_set(current.get(port_key, ''))
            added_ports = wanted_ports - current_ports
            if added_ports:
                if mode == 'WKA':
                    plan.append({'action': 'CONFLICT', 'name': name, 'group': group,
                                 'message': 'WKA ApplicationPort cannot be merged'})
                    continue
                changes[port_key] = _format_ports(current_ports | wanted_ports)
                labels.append('ADD_PORT')
        if changes:
            # A specific-application PUT must carry the complete approved desired
            # contract, not only the changed address/port. nGeniusONE v6.3.4 can
            # reject a sparse SBA payload when ApplicationType is absent. Start
            # from desired so type, parent and unchanged managed values remain
            # explicit, then overlay the additive unions calculated above.
            update_group = {'name': name, 'mode': mode,
                            'fields': {**desired, **changes}, 'rows': group['rows']}
            verify_group = {'name': name, 'mode': mode,
                            'fields': {**desired, **changes}, 'rows': group['rows']}
            plan.append({'action': '+'.join(labels), 'name': name, 'group': group,
                         'update_group': update_group, 'verify_group': verify_group,
                         'message': ', '.join(changes)})
        else:
            plan.append({'action': 'MATCHED', 'name': name, 'group': group, 'message': 'Already matches'})
    return plan
