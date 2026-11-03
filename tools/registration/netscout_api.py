"""Small NETSCOUT Config API client used by registration and sync tools."""
import base64
import ssl
import urllib.error
import urllib.parse
import urllib.request


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *unused):
        return None


class NetscoutClient:
    def __init__(self, base_url, username, password, ca_file=None, insecure=False, error_handler=None):
        parsed = urllib.parse.urlsplit(base_url)
        if (parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password
                or parsed.path not in ('', '/') or parsed.query or parsed.fragment):
            raise ValueError('--base-url must be https://host:port with no path, credentials or query')
        if ca_file and insecure:
            raise ValueError('--ca-file and --insecure cannot be used together')
        context = ssl._create_unverified_context() if insecure else ssl.create_default_context(cafile=ca_file)
        self.opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=context), NoRedirect())
        self.endpoint = base_url.rstrip('/') + '/ng1api/ncm/applications'
        self.auth = base64.b64encode(f'{username}:{password}'.encode()).decode()
        self.error_handler = error_handler

    def request(self, method, payload=None, application_name=None):
        url = self.endpoint
        if application_name is not None:
            url += '/' + urllib.parse.quote(application_name, safe='')
        req = urllib.request.Request(url, data=payload, method=method, headers={
            'Authorization': 'Basic ' + self.auth,
            'Accept': 'application/xml',
            'Content-Type': 'application/xml',
        })
        try:
            response = self.opener.open(req, timeout=30)
        except urllib.error.HTTPError as exc:
            detail = exc.read()
            if self.error_handler:
                self.error_handler(exc.code, detail)
            raise ValueError(f'HTTP {exc.code}') from exc
        with response:
            data = response.read()
            if response.status != 200:
                if self.error_handler:
                    self.error_handler(response.status, data)
                raise ValueError(f'Unexpected HTTP {response.status}; partial success is not accepted')
            return data

    def get_all(self):
        return self.request('GET')

    def create(self, payload):
        return self.request('POST', payload)

    def update(self, application_name, payload):
        return self.request('PUT', payload, application_name)
