"""IP 후보 YAML/Excel을 읽어 hostname 조회 사이트와 매칭한다.

`extract_ip_port_candidates.py`가 만든 `ip_port_candidates.yaml` 또는 Excel의
`Unique_IPs` 시트를 입력으로 사용한다. hostname 조회 사이트 URL과 HTML table
컬럼 위치는 별도 YAML 설정 파일로 받는다. API/NETSCOUT 호출은 하지 않는다.

지원 사이트 형식:

- `table`: HTML table의 td 텍스트를 hostname/ip 컬럼 index로 매칭한다.
- `input`: 같은 tr 안에 있는 input name=IP / input name=HOSTNAME 값을 매칭한다.

검색 결과에 유사 IP가 여러 개 있어도 target IP와 완전히 같은 IP만 채택한다.
"""

from __future__ import annotations

import argparse
import sys
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

try:
    import requests
    import yaml
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ModuleNotFoundError:
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    import requests
    import yaml
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font


UNIQUE_IP_SHEET = "Unique_IPs"
RESULT_COLUMNS = ["IP", "Hostname", "Applications", "Found", "Site", "Source URL", "Error"]
SUMMARY_COLUMNS = ["metric", "value"]


class HostnameLookupError(RuntimeError):
    """Hostname 조회 도구에서 복구할 수 없는 오류가 발생했을 때 사용한다."""


class TableParser(HTMLParser):
    """hostname 조회 사이트의 HTML table row를 단순 추출한다."""

    def __init__(self, require_td_align_center: bool = True) -> None:
        super().__init__()
        self.require_td_align_center = require_td_align_center
        self.in_tr = False
        self.in_td = False
        self.current_row: list[str] = []
        self.current_text = ""
        self.rows: list[list[str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() == "tr":
            self.in_tr = True
            self.current_row = []
        elif tag.lower() == "td" and self.in_tr:
            attrs_dict = {key: value or "" for key, value in attrs}
            if not self.require_td_align_center or attrs_dict.get("align", "").lower() == "center":
                self.in_td = True
                self.current_text = ""

    def handle_data(self, data: str) -> None:
        if self.in_td:
            self.current_text += data

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "td" and self.in_td:
            self.current_row.append(self.current_text.strip())
            self.current_text = ""
            self.in_td = False
        elif tag.lower() == "tr" and self.in_tr:
            if self.current_row:
                self.rows.append(self.current_row)
            self.current_row = []
            self.in_tr = False


class InputParser(HTMLParser):
    """같은 TR 안의 IP/HOSTNAME input value를 한 쌍으로 추출한다."""

    def __init__(self, require_td_align_center: bool = True) -> None:
        super().__init__()
        self.require_td_align_center = require_td_align_center
        self.in_tr = False
        self.in_target_td = False
        self.current_ip = ""
        self.current_hostname = ""
        self.results: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        attrs_dict = {key.lower(): value or "" for key, value in attrs}

        if tag == "tr":
            self.in_tr = True
            self.in_target_td = False
            self.current_ip = ""
            self.current_hostname = ""
            return

        if tag == "td" and self.in_tr:
            align = attrs_dict.get("align", "").lower()
            self.in_target_td = not self.require_td_align_center or align == "center"
            return

        if tag == "input" and self.in_tr and self.in_target_td:
            name = attrs_dict.get("name", "").strip().upper()
            value = attrs_dict.get("value", "").strip()
            if name == "IP":
                self.current_ip = value
            elif name == "HOSTNAME":
                self.current_hostname = value

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "td":
            self.in_target_td = False
            return
        if tag == "tr" and self.in_tr:
            if self.current_ip:
                self.results.append({"IP": self.current_ip, "HOSTNAME": self.current_hostname})
            self.in_tr = False
            self.in_target_td = False
            self.current_ip = ""
            self.current_hostname = ""


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise HostnameLookupError(f"YAML file not found: {path}")
    with path.open("r", encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    if not isinstance(data, dict):
        raise HostnameLookupError(f"YAML root must be a mapping: {path}")
    return data


def load_lookup_config(path: Path) -> dict[str, Any]:
    """hostname 조회 설정 YAML을 읽고 기본값을 채운다."""
    data = load_yaml(path)
    sites = data.get("sites", [])
    if not isinstance(sites, list) or not sites:
        raise HostnameLookupError("lookup config must contain non-empty sites list.")
    return {
        "sites": sites,
        "search_param_name": data.get("search_param_name", "ip"),
        "hostname_index": int(data.get("hostname_index", 1)),
        "ip_index": int(data.get("ip_index", 3)),
        "timeout": int(data.get("timeout", 10)),
        "require_td_align_center": bool(data.get("require_td_align_center", True)),
    }


def load_ip_applications_from_yaml(path: Path) -> dict[str, list[str]]:
    data = load_yaml(path)
    ips = data.get("unique_ip_ports", data.get("unique_ips", []))
    if not isinstance(ips, list):
        raise HostnameLookupError("ip candidates YAML field 'unique_ip_ports' must be a list.")
    result: dict[str, list[str]] = {}
    for item in ips:
        if isinstance(item, dict):
            ip = str(item.get("ip") or "").strip()
            applications = item.get("applications", [])
            if isinstance(applications, list):
                app_names = [str(app).strip() for app in applications if str(app or "").strip()]
            else:
                app_names = [str(applications).strip()] if str(applications or "").strip() else []
        else:
            ip = str(item or "").strip()
            app_names = []
        if not ip:
            continue
        if ip not in result:
            result[ip] = app_names
            continue
        merged = list(result[ip])
        for app in app_names:
            if app and app not in merged:
                merged.append(app)
        result[ip] = merged
    return result


def load_ip_applications_from_xlsx(path: Path) -> dict[str, list[str]]:
    if not path.exists():
        raise HostnameLookupError(f"XLSX file not found: {path}")
    workbook = load_workbook(path, read_only=True, data_only=True)
    if UNIQUE_IP_SHEET not in workbook.sheetnames:
        raise HostnameLookupError(f"Sheet not found: {UNIQUE_IP_SHEET} in {path}")
    worksheet = workbook[UNIQUE_IP_SHEET]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return {}
    headers = [str(header or "").strip() for header in rows[0]]
    try:
        ip_index = headers.index("IP")
    except ValueError as exc:
        raise HostnameLookupError(f"Column 'IP' not found in {UNIQUE_IP_SHEET}.") from exc
    applications_index = headers.index("Applications") if "Applications" in headers else None
    result: dict[str, list[str]] = {}
    for row in rows[1:]:
        if not row or not row[ip_index]:
            continue
        ip = str(row[ip_index]).strip()
        if not ip or ip in result:
            continue
        applications = []
        if applications_index is not None and applications_index < len(row) and row[applications_index]:
            applications = [app.strip() for app in str(row[applications_index]).split(",") if app.strip()]
        result[ip] = applications
    return result


def dedup_ips(values: Any) -> list[str]:
    seen = set()
    result = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def parse_html_rows(html: str, require_td_align_center: bool) -> list[list[str]]:
    parser = TableParser(require_td_align_center=require_td_align_center)
    parser.feed(html)
    return parser.rows


def find_exact_ip(rows: list[list[str]], target_ip: str, hostname_index: int, ip_index: int) -> dict[str, str] | None:
    required_index = max(hostname_index, ip_index)
    for row in rows:
        if len(row) <= required_index:
            continue
        hostname = row[hostname_index].strip()
        ip = row[ip_index].strip()
        if ip == target_ip:
            return {"ip": ip, "hostname": hostname}
    return None


def find_input_ip(html: str, target_ip: str, require_td_align_center: bool = True) -> dict[str, str] | None:
    """input 기반 조회 결과에서 target IP와 완전히 같은 항목만 반환한다."""
    parser = InputParser(require_td_align_center=require_td_align_center)
    parser.feed(html)
    parser.close()

    target_ip = target_ip.strip()
    for item in parser.results:
        ip = str(item.get("IP", "")).strip()
        hostname = str(item.get("HOSTNAME", "")).strip()
        if ip == target_ip:
            return {"ip": ip, "hostname": hostname}
    return None


def request_site(url: str, target_ip: str, search_param_name: str, timeout: int) -> str:
    response = requests.get(url, params={search_param_name: target_ip}, timeout=timeout)
    response.raise_for_status()
    return response.text


def lookup_one_ip(target_ip: str, applications: list[str], config: dict[str, Any]) -> dict[str, Any]:
    """IP 하나를 설정된 사이트 순서대로 조회하고 첫 매칭을 반환한다."""
    errors = []
    for index, site in enumerate(config["sites"], start=1):
        if isinstance(site, str):
            site_name = f"Site{index}"
            url = site
            site_type = "table"
            search_param_name = config["search_param_name"]
            hostname_index = config["hostname_index"]
            ip_index = config["ip_index"]
            require_center = config["require_td_align_center"]
        else:
            site_name = str(site.get("name") or f"Site{index}")
            url = str(site.get("url") or "")
            site_type = str(site.get("type") or "table").strip().lower()
            search_param_name = str(site.get("search_param_name") or config["search_param_name"])
            hostname_index = int(site.get("hostname_index", config["hostname_index"]))
            ip_index = int(site.get("ip_index", config["ip_index"]))
            require_center = bool(site.get("require_td_align_center", config["require_td_align_center"]))
        if not url:
            continue
        try:
            html = request_site(url, target_ip, search_param_name, config["timeout"])
            if site_type == "table":
                rows = parse_html_rows(html, require_center)
                found = find_exact_ip(rows, target_ip, hostname_index, ip_index)
            elif site_type == "input":
                found = find_input_ip(html, target_ip, require_center)
            else:
                raise HostnameLookupError(f"Unknown site type '{site_type}' for {site_name}")
            if found:
                return {
                    "IP": target_ip,
                    "Hostname": found["hostname"],
                    "Applications": ", ".join(applications),
                    "Found": True,
                    "Site": site_name,
                    "Source URL": url,
                    "Error": "",
                }
        except requests.exceptions.RequestException as exc:
            errors.append(f"{site_name}: {exc}")
        except HostnameLookupError as exc:
            errors.append(f"{site_name}: {exc}")
    return {
        "IP": target_ip,
        "Hostname": "",
        "Applications": ", ".join(applications),
        "Found": False,
        "Site": "",
        "Source URL": "",
        "Error": " | ".join(errors),
    }


def write_results_xlsx(results: list[dict[str, Any]], output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Hostname_Lookup"
    worksheet.append(RESULT_COLUMNS)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    for row in results:
        worksheet.append([row.get(column) for column in RESULT_COLUMNS])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_cells in worksheet.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 80)

    summary = workbook.create_sheet("Run_Summary")
    summary.append(SUMMARY_COLUMNS)
    found_count = sum(1 for row in results if row.get("Found"))
    for row in (
        {"metric": "ip_count", "value": len(results)},
        {"metric": "found_count", "value": found_count},
        {"metric": "not_found_count", "value": len(results) - found_count},
    ):
        summary.append([row["metric"], row["value"]])
    workbook.save(output_path)
    return output_path


def lookup_hostnames(
    config_path: Path,
    output_path: Path,
    ip_candidates_yaml: Path | None = None,
    ip_candidates_xlsx: Path | None = None,
) -> tuple[Path, int, int]:
    if not ip_candidates_yaml and not ip_candidates_xlsx:
        raise HostnameLookupError("Provide --ip-candidates-yaml or --ip-candidates-xlsx.")
    config = load_lookup_config(config_path)
    ip_applications = (
        load_ip_applications_from_yaml(ip_candidates_yaml)
        if ip_candidates_yaml
        else load_ip_applications_from_xlsx(ip_candidates_xlsx)  # type: ignore[arg-type]
    )
    results = [lookup_one_ip(ip, applications, config) for ip, applications in ip_applications.items()]
    path = write_results_xlsx(results, output_path)
    found_count = sum(1 for row in results if row.get("Found"))
    return path, len(results), found_count


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Lookup hostnames for IP candidates exported from NETSCOUT Detail results.")
    parser.add_argument("--config", required=True, type=Path, help="hostname lookup site 설정 YAML 경로.")
    parser.add_argument("--ip-candidates-yaml", type=Path, default=None, help="extract_ip_port_candidates.py가 만든 YAML 경로.")
    parser.add_argument("--ip-candidates-xlsx", type=Path, default=None, help="extract_ip_port_candidates.py가 만든 Excel 경로.")
    parser.add_argument("--output", required=True, type=Path, help="hostname 조회 결과 Excel 경로.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        path, ip_count, found_count = lookup_hostnames(args.config, args.output, args.ip_candidates_yaml, args.ip_candidates_xlsx)
    except HostnameLookupError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(f"completed ip_count={ip_count} found_count={found_count} not_found_count={ip_count - found_count} output={path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
