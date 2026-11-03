"""고객사 hostname 조회 사이트를 흉내내는 HTTPS lab 서버.

하나의 컨테이너에서 두 포트를 동시에 띄운다.

- 5000: HTML table 응답
- 5002: input value 기반 응답

NETSCOUT API mock이 아니다. NETSCOUT 조회는 실제 테스트 서버를 사용하고,
이 lab은 `ip_port_candidates.yaml` 이후 hostname 조회 단계만 검증한다.
"""

from __future__ import annotations

import html
import ssl
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse


HOSTS = {
    "192.0.2.10": "md-feed-gw-01",
    "192.0.2.10": "md-feed-gw-02",
    "192.0.2.10": "md-feed-cache-01",
    "192.0.2.10": "trading-app-01",
    "192.0.2.10": "trading-app-02",
    "192.0.2.10": "dns-core-01",
    "192.0.2.10": "dns-core-02",
    "192.0.2.10": "netscout-test-server",
    "192.0.2.10": "monitor-if3-host",
    "192.0.2.10": "market-data-router-01",
    "192.0.2.10": "market-data-router-10",
    "192.0.2.10": "market-data-router-12",
    "192.0.2.10": "market-data-router-100",
    "192.0.2.10": "order-gw-01",
    "192.0.2.10": "order-gw-02",
}


class HostnameHandler(BaseHTTPRequestHandler):
    response_type = "table"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/search":
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"not found")
            return

        query = parse_qs(parsed.query)
        target_ip = (query.get("search") or query.get("ip") or [""])[0].strip()
        body = self.render_input(target_ip) if self.response_type == "input" else self.render_table(target_ip)
        payload = body.encode("utf-8")

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def log_message(self, format: str, *args) -> None:
        print(f"[{self.response_type}] {self.address_string()} {format % args}")

    def matched_rows(self, target_ip: str) -> list[tuple[str, str]]:
        rows = list(HOSTS.items())
        if target_ip in HOSTS:
            rows.insert(0, (target_ip, HOSTS[target_ip]))
        elif target_ip:
            rows.insert(0, (target_ip, ""))
        return rows

    def render_table(self, target_ip: str) -> str:
        rows = []
        for ip, hostname in self.matched_rows(target_ip):
            rows.append(
                "<tr>"
                f"<td>{html.escape(hostname)}</td>"
                f"<td>{html.escape(ip)}</td>"
                f"<td>lab</td>"
                "</tr>"
            )
        return (
            "<html><body>"
            "<h1>Hostname Lab Site1</h1>"
            "<table>"
            "<tr><th>Hostname</th><th>IP</th><th>Source</th></tr>"
            + "".join(rows)
            + "</table></body></html>"
        )

    def render_input(self, target_ip: str) -> str:
        rows = []
        for ip, hostname in self.matched_rows(target_ip):
            rows.append(
                "<tr><td>"
                f'<input name="IP" value="{html.escape(ip)}">'
                f'<input name="HOSTNAME" value="{html.escape(hostname)}">'
                "</td></tr>"
            )
        return (
            "<html><body>"
            "<h1>Hostname Lab Site2</h1>"
            "<table>"
            + "".join(rows)
            + "</table></body></html>"
        )


def make_handler(response_type: str):
    return type(f"{response_type.title()}HostnameHandler", (HostnameHandler,), {"response_type": response_type})


def serve(port: int, response_type: str) -> None:
    server = ThreadingHTTPServer(("0.0.0.0", port), make_handler(response_type))
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.load_cert_chain("/app/server.crt", "/app/server.key")
    server.socket = context.wrap_socket(server.socket, server_side=True)
    print(f"hostname lab {response_type} server listening on https://0.0.0.0:{port}/search")
    server.serve_forever()


def main() -> None:
    threads = [
        threading.Thread(target=serve, args=(5000, "table"), daemon=True),
        threading.Thread(target=serve, args=(5002, "input"), daemon=True),
    ]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()


if __name__ == "__main__":
    main()
