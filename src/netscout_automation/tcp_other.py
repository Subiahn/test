"""TCP_OTHER Search API, Data API query, 정규화, XLSX export 로직.

중요한 NETSCOUT 데이터 모델:
- Search API의 `ipAddress`/`ifn`은 검색 대상 Host 자체가 아니라 관측 지점일 수 있다.
- Monitor IP / Interface No.는 트래픽 관측 지점이다.
- TOPNCONV의 `srcAddress`/`destAddress`/port는 실제 conversation endpoint이다.
- TCP_OTHER는 Config API Application List에서 확인된 application_type_code/application_short_name이다.
- APPLICATION query는 Port/Trx를 제공하지 않아 Byte/Packet 중심으로 사용한다.
- TOPNCONV query는 Port/Byte/Packet 확인용으로 사용한다.
- Trx는 현재 응답에서 확인되지 않아 TODO로 남긴다.
"""

from __future__ import annotations

import json
import logging
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from time import perf_counter
from urllib.parse import quote, urljoin, urlparse
from xml.sax.saxutils import escape

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font
from requests.auth import HTTPBasicAuth

from .config import NetscoutSettings

logger = logging.getLogger(__name__)

TCP_OTHER_REPORT_COLUMNS = [
    "IP",
    "Interface No.",
    "App ID",
    "Application",
    "Byte",
    "Byte In",
    "Byte Out",
    "Packets",
    "Packets In",
    "Packets Out",
]

TCP_OTHER_TOPNCONV_REPORT_COLUMNS = [
    "Application",
    "App ID",
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "Byte",
    "Byte In",
    "Byte Out",
    "Packets",
    "Packets In",
    "Packets Out",
    "Monitor IP",
    "Interface No.",
]

TCP_OTHER_HOST_SUMMARY_COLUMNS = [
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "Application",
    "Total Byte",
    "Total Byte In",
    "Total Byte Out",
    "Total Packets",
    "Total Packets In",
    "Total Packets Out",
    "Observed Interfaces",
    "Observed Interface Count",
]

TCP_OTHER_REPORT_MODES = {"application", "topnconv"}


class TcpOtherError(RuntimeError):
    """TCP_OTHER report 처리 단계가 실패했을 때 발생한다."""


@dataclass(frozen=True)
class SearchTarget:
    """Data API Search가 반환한 monitor 관측 지점. conversation endpoint가 아니다."""

    ip_address: str
    ifn: str


@dataclass(frozen=True)
class TcpOtherReportResult:
    searched_targets: int
    detail_success_count: int
    detail_failure_count: int
    row_count: int
    output_path: Path


def search_application(settings: NetscoutSettings, application: str) -> list[SearchTarget]:
    """특정 application이 관측된 monitor IP/Interface 조합을 Search API로 찾는다."""
    endpoint = _search_endpoint(settings.data_api_endpoint, application)
    url = _url(settings.base_url, endpoint)
    host = urlparse(url).hostname or "unknown-host"
    logger.info("Calling Data API Search host=%s endpoint=%s verify_ssl=%s", host, endpoint, settings.verify_ssl)

    try:
        if not settings.verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        response = requests.get(
            url,
            auth=HTTPBasicAuth(settings.username, settings.password),
            headers={"Content-Type": "application/xml"},
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
    except requests.exceptions.Timeout as exc:
        raise TcpOtherError("Data API Search timed out. Check network and timeout settings.") from exc
    except requests.exceptions.SSLError as exc:
        raise TcpOtherError("Data API Search failed with an SSL error. Check verify_ssl and server certificate.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise TcpOtherError("Data API Search connection failed. Check base_url and network access.") from exc
    except requests.exceptions.RequestException as exc:
        raise TcpOtherError("Data API Search request failed.") from exc

    _raise_for_status(response, "Data API Search")
    targets = parse_search_application_response(response.text)
    logger.info("Data API Search returned %s IP/Interface targets", len(targets))
    return targets


def parse_search_application_response(xml_text: str) -> list[SearchTarget]:
    """Search API XML을 파싱하고 중복 없는 monitor IP/interface 조합만 남긴다."""
    if not xml_text.strip():
        return []
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as exc:
        raise TcpOtherError("Data API Search returned invalid XML.") from exc

    targets: list[SearchTarget] = []
    for row in root.findall(".//outputRow"):
        ip_address = _child_text(row, "ipAddress")
        ifn = _child_text(row, "ifn")
        if ip_address and ifn:
            targets.append(SearchTarget(ip_address=ip_address, ifn=ifn))
    return deduplicate_search_targets(targets)


def deduplicate_search_targets(targets: list[SearchTarget]) -> list[SearchTarget]:
    """Search API 반환 순서를 유지하면서 중복 monitor 관측 지점을 제거한다."""
    deduplicated: list[SearchTarget] = []
    seen: set[tuple[str, str]] = set()
    for target in targets:
        key = (target.ip_address, target.ifn)
        if key in seen:
            continue
        seen.add(key)
        deduplicated.append(target)
    return deduplicated


def build_tcp_other_query_xml(ip: str, ifn: str, duration: str) -> str:
    """하나의 IP/Interface 조합에 대한 APPLICATION table query XML을 만든다."""
    return f"""<GenericClientQuery>
  <NetworkObjectData>
    <NetworkParameter>APPLICATION</NetworkParameter>
    <SelectColumnList>
      <ClientColumn>ipAddress</ClientColumn>
      <ClientColumn>ifn</ClientColumn>
      <ClientColumn>appId</ClientColumn>
      <ClientColumn>appId_String</ClientColumn>
      <ClientColumn>octets</ClientColumn>
      <ClientColumn>octetsIn</ClientColumn>
      <ClientColumn>octetsOut</ClientColumn>
      <ClientColumn>packets</ClientColumn>
      <ClientColumn>packetsIn</ClientColumn>
      <ClientColumn>packetsOut</ClientColumn>
    </SelectColumnList>
  </NetworkObjectData>
  <FlowFilterList>
    <FlowFilter>
      <IPAddress>{escape(ip)}</IPAddress>
      <Ifn>{escape(ifn)}</Ifn>
      <FilterList />
    </FlowFilter>
  </FlowFilterList>
  <FunctionList />
  <TimeDef>
    <duration>{escape(duration)}</duration>
    <resolution>NO_RESOLUTION</resolution>
  </TimeDef>
  <params>
    <columnOrder>true</columnOrder>
    <stringConversion>true</stringConversion>
  </params>
</GenericClientQuery>"""


def build_tcp_other_topnconv_query_xml(ip: str, ifn: str, duration: str) -> str:
    """하나의 IP/Interface 조합에 대한 TOPNCONV query XML을 만든다."""
    return f"""<GenericClientQuery>
  <NetworkObjectData>
    <NetworkParameter>TOPNCONV</NetworkParameter>
    <SelectColumnList>
      <ClientColumn>appId</ClientColumn>
      <ClientColumn>appId_String</ClientColumn>
      <ClientColumn>srcAddress</ClientColumn>
      <ClientColumn>destAddress</ClientColumn>
      <ClientColumn>srcPort</ClientColumn>
      <ClientColumn>destPort</ClientColumn>
      <ClientColumn>srcToDestOctets</ClientColumn>
      <ClientColumn>destToSrcOctets</ClientColumn>
      <ClientColumn>srcToDestPackets</ClientColumn>
      <ClientColumn>destToSrcPackets</ClientColumn>
      <ClientColumn>octets</ClientColumn>
      <ClientColumn>packets</ClientColumn>
    </SelectColumnList>
  </NetworkObjectData>
  <FlowFilterList>
    <FlowFilter>
      <IPAddress>{escape(ip)}</IPAddress>
      <Ifn>{escape(ifn)}</Ifn>
      <!-- 서버사이드 appId 필터는 probe 중 불안정할 수 있어 응답 후 Python에서 다시 필터링한다. -->
      <FilterList />
    </FlowFilter>
  </FlowFilterList>
  <FunctionList>
    <Function>
      <name>TopN</name>
      <column>octets</column>
      <order>descending</order>
      <nValue>100</nValue>
    </Function>
  </FunctionList>
  <TimeDef>
    <duration>{escape(duration)}</duration>
    <resolution>NO_RESOLUTION</resolution>
  </TimeDef>
  <params>
    <columnOrder>true</columnOrder>
    <stringConversion>true</stringConversion>
  </params>
</GenericClientQuery>"""


def fetch_tcp_other_detail(
    settings: NetscoutSettings,
    ip: str,
    ifn: str,
    application: str,
    duration: str,
) -> list[dict[str, object]]:
    """APPLICATION 데이터를 조회하고 appId_String이 일치하는 row만 남긴다."""
    return _fetch_tcp_other_rows(
        settings=settings,
        ip=ip,
        ifn=ifn,
        application=application,
        duration=duration,
        query_xml=build_tcp_other_query_xml(ip, ifn, duration),
        label="Data API APPLICATION detail",
        include_monitor_fields=False,
    )


def fetch_tcp_other_topnconv_detail(
    settings: NetscoutSettings,
    ip: str,
    ifn: str,
    application: str,
    duration: str,
) -> list[dict[str, object]]:
    """TOPNCONV 데이터를 조회하고 appId_String이 일치하는 row만 남긴다."""
    return _fetch_tcp_other_rows(
        settings=settings,
        ip=ip,
        ifn=ifn,
        application=application,
        duration=duration,
        query_xml=build_tcp_other_topnconv_query_xml(ip, ifn, duration),
        label="Data API TOPNCONV detail",
        include_monitor_fields=True,
    )


def _fetch_tcp_other_rows(
    settings: NetscoutSettings,
    ip: str,
    ifn: str,
    application: str,
    duration: str,
    query_xml: str,
    label: str,
    include_monitor_fields: bool,
) -> list[dict[str, object]]:
    url = _url(settings.base_url, settings.data_api_endpoint)
    host = urlparse(url).hostname or "unknown-host"
    logger.info(
        "Calling %s host=%s endpoint=%s ip=%s ifn=%s verify_ssl=%s",
        label,
        host,
        settings.data_api_endpoint,
        ip,
        ifn,
        settings.verify_ssl,
    )

    try:
        if not settings.verify_ssl:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        start = perf_counter()
        response = requests.post(
            url,
            auth=HTTPBasicAuth(settings.username, settings.password),
            headers={"Content-Type": "application/xml", "Accept": "application/json"},
            data=query_xml.encode("utf-8"),
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
        elapsed = perf_counter() - start
    except requests.exceptions.Timeout as exc:
        raise TcpOtherError(f"{label} timed out for ip={ip} ifn={ifn}.") from exc
    except requests.exceptions.SSLError as exc:
        raise TcpOtherError(f"{label} SSL error for ip={ip} ifn={ifn}.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise TcpOtherError(f"{label} connection failed for ip={ip} ifn={ifn}.") from exc
    except requests.exceptions.RequestException as exc:
        raise TcpOtherError(f"{label} request failed for ip={ip} ifn={ifn}.") from exc

    logger.info(
        "%s response status=%s elapsed=%.3fs content_type=%s ip=%s ifn=%s",
        label,
        response.status_code,
        elapsed,
        response.headers.get("Content-Type", "unknown"),
        ip,
        ifn,
    )
    _raise_for_status(response, label)

    try:
        payload = response.json()
    except ValueError as exc:
        raise TcpOtherError(f"{label} returned invalid JSON for ip={ip} ifn={ifn}.") from exc

    rows = payload.get("response", [])
    if not isinstance(rows, list):
        raise TcpOtherError(f"{label} JSON response field is not a list for ip={ip} ifn={ifn}.")

    expected = application.upper()
    # XML에 FilterList가 있어도 Python-side 필터를 유지한다. NETSCOUT이 주변 app row를 함께 줄 수 있다.
    filtered_rows = [row for row in rows if isinstance(row, dict) and str(row.get("appId_String", "")).upper() == expected]
    if include_monitor_fields:
        # Monitor IP / Interface No.는 src/dest conversation endpoint와 별개로 보존한다.
        for row in filtered_rows:
            row.setdefault("_monitor_ip", ip)
            row.setdefault("_ifn", ifn)
    return filtered_rows


def normalize_tcp_other_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    normalized: list[dict[str, object]] = []
    for row in rows:
        normalized.append(
            {
                "IP": row.get("ipAddress"),
                "Interface No.": row.get("ifn"),
                "App ID": row.get("appId"),
                "Application": row.get("appId_String"),
                "Byte": row.get("octets"),
                "Byte In": row.get("octetsIn"),
                "Byte Out": row.get("octetsOut"),
                "Packets": row.get("packets"),
                "Packets In": row.get("packetsIn"),
                "Packets Out": row.get("packetsOut"),
            }
        )
    return normalized


def normalize_tcp_other_topnconv_rows(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """TOPNCONV 응답 field를 검토용 column으로 매핑하고 Byte 내림차순으로 정렬한다."""
    normalized: list[dict[str, object]] = []
    for row in rows:
        normalized.append(
            {
                "Application": row.get("appId_String"),
                "App ID": row.get("appId"),
                "Source IP": row.get("srcAddress"),
                "Source Port": row.get("srcPort"),
                "Destination IP": row.get("destAddress"),
                "Destination Port": row.get("destPort"),
                "Byte": row.get("octets"),
                "Byte In": row.get("srcToDestOctets"),
                "Byte Out": row.get("destToSrcOctets"),
                "Packets": row.get("packets"),
                "Packets In": row.get("srcToDestPackets"),
                "Packets Out": row.get("destToSrcPackets"),
                "Monitor IP": row.get("_monitor_ip"),
                "Interface No.": row.get("_ifn"),
            }
        )
    return sorted(normalized, key=lambda row: _numeric_sort_value(row.get("Byte")), reverse=True)


def summarize_tcp_other_by_host(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """Source/Destination/Port/Application 기준으로 TCP_OTHER conversation을 집계한다."""
    grouped: dict[tuple[object, object, object, object, object], dict[str, object]] = {}
    observed: dict[tuple[object, object, object, object, object], set[str]] = {}

    for row in rows:
        key = (
            row.get("Source IP"),
            row.get("Source Port"),
            row.get("Destination IP"),
            row.get("Destination Port"),
            row.get("Application"),
        )
        if key not in grouped:
            grouped[key] = {
                "Source IP": row.get("Source IP"),
                "Source Port": row.get("Source Port"),
                "Destination IP": row.get("Destination IP"),
                "Destination Port": row.get("Destination Port"),
                "Application": row.get("Application"),
                "Total Byte": 0,
                "Total Byte In": 0,
                "Total Byte Out": 0,
                "Total Packets": 0,
                "Total Packets In": 0,
                "Total Packets Out": 0,
                "Observed Interfaces": "",
                "Observed Interface Count": 0,
            }
            observed[key] = set()

        grouped[key]["Total Byte"] = _sum_number(grouped[key]["Total Byte"], row.get("Byte"))
        grouped[key]["Total Byte In"] = _sum_number(grouped[key]["Total Byte In"], row.get("Byte In"))
        grouped[key]["Total Byte Out"] = _sum_number(grouped[key]["Total Byte Out"], row.get("Byte Out"))
        grouped[key]["Total Packets"] = _sum_number(grouped[key]["Total Packets"], row.get("Packets"))
        grouped[key]["Total Packets In"] = _sum_number(grouped[key]["Total Packets In"], row.get("Packets In"))
        grouped[key]["Total Packets Out"] = _sum_number(grouped[key]["Total Packets Out"], row.get("Packets Out"))

        monitor_ip = row.get("Monitor IP")
        interface_no = row.get("Interface No.")
        if monitor_ip not in (None, "") and interface_no not in (None, ""):
            observed[key].add(f"{monitor_ip}|{interface_no}")

    summaries = []
    for key, summary in grouped.items():
        interfaces = sorted(observed[key])
        summary["Observed Interfaces"] = ", ".join(interfaces)
        summary["Observed Interface Count"] = len(interfaces)
        summaries.append(summary)
    return sorted(summaries, key=lambda row: _numeric_sort_value(row.get("Total Byte")), reverse=True)


def export_tcp_other_report_xlsx(
    rows: list[dict[str, object]],
    output_path: str | Path,
    columns: list[str] | None = None,
) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    selected_columns = columns or TCP_OTHER_REPORT_COLUMNS

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "TCP_OTHER"
    worksheet.append(selected_columns)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for row in rows:
        worksheet.append([row.get(column) for column in selected_columns])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_cells in worksheet.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 48)

    workbook.save(path)
    return path


def export_tcp_other_report_with_summary_xlsx(
    detail_rows: list[dict[str, object]],
    summary_rows: list[dict[str, object]],
    output_path: str | Path,
) -> Path:
    """TCP_OTHER 상세와 Source/Destination host 요약을 같은 workbook에 저장한다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    _write_sheet(workbook.active, "TCP_OTHER_Detail", TCP_OTHER_TOPNCONV_REPORT_COLUMNS, detail_rows)
    _write_sheet(workbook.create_sheet("TCP_OTHER_Host_Summary"), "TCP_OTHER_Host_Summary", TCP_OTHER_HOST_SUMMARY_COLUMNS, summary_rows)
    workbook.save(path)
    return path


def generate_tcp_other_report(
    settings: NetscoutSettings,
    application: str,
    duration: str,
    output_path: str | Path,
    mode: str = "topnconv",
    summary_by_host: bool = False,
) -> TcpOtherReportResult:
    """APPLICATION legacy mode 또는 TOPNCONV detail mode로 TCP_OTHER report를 생성한다."""
    mode = _validate_report_mode(mode)
    targets = search_application(settings, application)
    all_rows: list[dict[str, object]] = []
    success_count = 0
    failure_count = 0

    for index, target in enumerate(targets, start=1):
        try:
            logger.info(
                "Fetching TCP_OTHER detail mode=%s target=%s/%s ip=%s ifn=%s",
                mode,
                index,
                len(targets),
                target.ip_address,
                target.ifn,
            )
            if mode == "topnconv":
                rows = fetch_tcp_other_topnconv_detail(settings, target.ip_address, target.ifn, application, duration)
            else:
                rows = fetch_tcp_other_detail(settings, target.ip_address, target.ifn, application, duration)
            all_rows.extend(rows)
            success_count += 1
        except TcpOtherError as exc:
            failure_count += 1
            logger.warning("Skipping failed Data API detail target ip=%s ifn=%s: %s", target.ip_address, target.ifn, exc)

    if mode == "topnconv":
        normalized_rows = normalize_tcp_other_topnconv_rows(all_rows)
        columns = TCP_OTHER_TOPNCONV_REPORT_COLUMNS
        if summary_by_host:
            summary_rows = summarize_tcp_other_by_host(normalized_rows)
            path = export_tcp_other_report_with_summary_xlsx(normalized_rows, summary_rows, output_path)
        else:
            path = export_tcp_other_report_xlsx(normalized_rows, output_path, columns)
    else:
        normalized_rows = normalize_tcp_other_rows(all_rows)
        columns = TCP_OTHER_REPORT_COLUMNS
        if summary_by_host:
            logger.warning("summary-by-host is only supported for topnconv mode; exporting detail sheet only")
        path = export_tcp_other_report_xlsx(normalized_rows, output_path, columns)
    logger.info(
        "TCP_OTHER report completed targets=%s succeeded=%s failed=%s rows=%s output=%s",
        len(targets),
        success_count,
        failure_count,
        len(normalized_rows),
        path,
    )
    return TcpOtherReportResult(
        searched_targets=len(targets),
        detail_success_count=success_count,
        detail_failure_count=failure_count,
        row_count=len(normalized_rows),
        output_path=path,
    )


def _search_endpoint(data_api_endpoint: str, application: str) -> str:
    base = data_api_endpoint.rstrip("/")
    return f"{base}/search/application/value/{quote(application, safe='')}"


def _validate_report_mode(mode: str) -> str:
    normalized = mode.lower()
    if normalized not in TCP_OTHER_REPORT_MODES:
        raise TcpOtherError(f"Unsupported tcp-other-report mode: {mode}")
    return normalized


def _numeric_sort_value(value: object) -> float:
    if value in (None, ""):
        return 0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0


def _sum_number(left: object, right: object) -> float:
    return _numeric_sort_value(left) + _numeric_sort_value(right)


def _write_sheet(worksheet, title: str, columns: list[str], rows: list[dict[str, object]]) -> None:
    worksheet.title = title
    worksheet.append(columns)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        worksheet.append([row.get(column) for column in columns])
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for column_cells in worksheet.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 48)


def _url(base_url: str, endpoint: str) -> str:
    return urljoin(base_url.rstrip("/") + "/", endpoint.lstrip("/"))


def _child_text(element: ET.Element, child_name: str) -> str:
    child = element.find(child_name)
    return (child.text or "").strip() if child is not None else ""


def _raise_for_status(response: requests.Response, action: str) -> None:
    if response.status_code == 200:
        return
    if response.status_code == 500:
        raise TcpOtherError(f"{action} failed with HTTP 500 internal server error.")
    preview = _sanitized_preview(response.text)
    raise TcpOtherError(f"{action} failed with HTTP {response.status_code}. Response preview={preview!r}")


def _sanitized_preview(text: str, limit: int = 300) -> str:
    compact = " ".join((text or "").split())
    # 인증정보가 포함될 수 있는 구조화된 error payload를 그대로 출력하지 않는다.
    try:
        parsed = json.loads(compact)
    except json.JSONDecodeError:
        return compact[:limit]
    if isinstance(parsed, dict):
        for key in list(parsed):
            if key.lower() in {"password", "token", "cookie", "authorization"}:
                parsed[key] = "***"
    return json.dumps(parsed, ensure_ascii=True)[:limit]

