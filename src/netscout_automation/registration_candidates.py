"""Application 등록 후보 분석 report 생성 로직."""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from time import perf_counter
from typing import Any
from urllib.parse import urljoin, urlparse
from xml.sax.saxutils import escape

import requests
import urllib3
from openpyxl import Workbook
from openpyxl.styles import Font
from requests.auth import HTTPBasicAuth

from .client import NetscoutClient, NetscoutClientError
from .config import NetscoutSettings
from .parsers import (
    APPLICATION_COLUMNS,
    DEVICE_COLUMNS,
    INTERFACE_COLUMNS,
    ParserError,
    parse_applications_response,
    parse_device_interfaces_response,
    parse_devices_response,
)
from .tcp_other import SearchTarget, TcpOtherError, deduplicate_search_targets, search_application

logger = logging.getLogger(__name__)

DEVICE_LIST_COLUMNS = DEVICE_COLUMNS
INTERFACE_LIST_COLUMNS = INTERFACE_COLUMNS
TOPNCONV_RAW_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "interfaceName",
    "interfaceNumber",
    "interface_key",
    "appId",
    "appId_String",
    "srcAddress",
    "srcPort",
    "destAddress",
    "destPort",
    "srcToDestOctets",
    "destToSrcOctets",
    "octets",
    "srcToDestPackets",
    "destToSrcPackets",
    "packets",
    "bps",
]
INTERFACE_TOPN_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "interfaceName",
    "interfaceNumber",
    "interface_key",
    "rank",
    "appId_String",
    "srcAddress",
    "srcPort",
    "destAddress",
    "destPort",
    "octets",
    "packets",
    "bps",
]
REGISTRATION_CANDIDATE_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "interfaceName",
    "interfaceNumber",
    "interface_key",
    "candidate_application_name",
    "candidate_port",
    "candidate_port_type",
    "protocol",
    "srcAddress",
    "srcPort",
    "destAddress",
    "destPort",
    "octets",
    "packets",
    "bps",
    "rank",
    "registration_reason",
]
SUMMARY_COLUMNS = ["metric", "value"]

DURATION_SECONDS = {
    "LAST_HOUR": 3600,
    "LAST_6_HOURS": 21600,
    "LAST_12_HOURS": 43200,
    "LAST_24_HOURS": 86400,
    "LAST_DAY": 86400,
    "LAST_7_DAYS": 604800,
    "LAST_WEEK": 604800,
}


class RegistrationCandidatesError(RuntimeError):
    """Application 등록 후보 report를 만들 수 없을 때 발생한다."""


@dataclass(frozen=True)
class RegistrationCandidatesResult:
    devices_count: int
    interfaces_count: int
    topn_rows_count: int
    tcp_other_rows_count: int
    registration_candidate_count: int
    output_path: Path


@dataclass(frozen=True)
class RegistrationDryRunPreview:
    endpoints: list[str]
    duration: str
    top_n: int
    output_path: Path
    credentials_present: bool


def build_interface_key(monitor_ip: object, interface_no: object) -> str:
    """InterfaceNumber 단독이 아니라 deviceIPAddress와 묶어 interface 식별자를 만든다."""
    return f"{monitor_ip}|{interface_no}"


def duration_to_seconds(duration: str) -> int | None:
    """Data API duration 값을 bps 계산용 초 단위로 변환한다."""
    return DURATION_SECONDS.get(duration.upper())


def build_topnconv_query_xml(ip: str, ifn: str, duration: str, top_n: int) -> str:
    """하나의 monitor IP/Interface에 대한 TOPNCONV query XML을 만든다."""
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
      <FilterList />
    </FlowFilter>
  </FlowFilterList>
  <FunctionList>
    <Function>
      <name>TopN</name>
      <column>octets</column>
      <order>descending</order>
      <nValue>{top_n}</nValue>
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


def build_dry_run_preview(settings: NetscoutSettings, duration: str, top_n: int, output_path: str | Path) -> RegistrationDryRunPreview:
    """실제 API 호출 없이 registration-candidates 실행 계획만 만든다."""
    return RegistrationDryRunPreview(
        endpoints=[
            settings.config_applications_endpoint,
            settings.devices_endpoint,
            settings.device_interfaces_endpoint_template,
            settings.data_api_endpoint,
        ],
        duration=duration,
        top_n=top_n,
        output_path=Path(output_path),
        credentials_present=bool(settings.username and settings.password),
    )


def discover_interfaces_with_search_api(settings: NetscoutSettings, application: str = "TCP_OTHER") -> list[dict[str, object]]:
    """Device/Interface API 미확인 상태에서 Search API 결과로 monitor interface를 발견한다."""
    targets = search_application(settings, application)
    return build_interface_rows_from_targets(targets)


def collect_devices(settings: NetscoutSettings) -> list[dict[str, Any]]:
    """Config API Device List를 수집하고 민감 field를 제외한 정규화 row로 변환한다."""
    return parse_devices_response(NetscoutClient(settings).get_devices())


def collect_interfaces_for_devices(
    settings: NetscoutSettings,
    devices: list[dict[str, Any]],
    active_only: bool = True,
    max_devices: int | None = None,
    max_interfaces: int | None = None,
) -> list[dict[str, Any]]:
    """Device별 Interface List를 수집하고 InterfaceNumber가 유효한 row만 남긴다."""
    selected_devices = [device for device in devices if not active_only or is_active_status(device.get("status"))]
    if max_devices is not None:
        selected_devices = selected_devices[:max_devices]

    interfaces: list[dict[str, Any]] = []
    for device in selected_devices:
        device_name = str(device.get("deviceName") or "")
        device_ip = str(device.get("deviceIPAddress") or "")
        if not device_name:
            logger.warning("Skipping device without deviceName device_ip=%s", device_ip)
            continue
        try:
            raw_response = NetscoutClient(settings).get_device_interfaces(device_name)
            device_interfaces = parse_device_interfaces_response(raw_response, device_name=device_name, device_ip_address=device_ip)
        except (NetscoutClientError, ParserError) as exc:
            logger.warning("Skipping failed Device Interface List device=%s: %s", device_name, exc)
            continue
        for interface in device_interfaces:
            if active_only and not is_active_status(interface.get("status")):
                continue
            if not is_valid_interface_number(interface.get("interfaceNumber")):
                logger.warning("Skipping interface with invalid InterfaceNumber device=%s value=%s", device_name, interface.get("interfaceNumber"))
                continue
            interfaces.append(interface)
            if max_interfaces is not None and len(interfaces) >= max_interfaces:
                return interfaces
    return interfaces


def is_active_status(value: object) -> bool:
    """Active/ACT 상태만 기본 처리 대상으로 본다. 값이 비어 있으면 보수적으로 제외한다."""
    return str(value or "").strip().lower() in {"active", "act"}


def is_valid_interface_number(value: object) -> bool:
    """Data API Ifn으로 사용할 수 있는 InterfaceNumber인지 확인한다."""
    if value in (None, ""):
        return False
    try:
        return int(value) >= 0
    except (TypeError, ValueError):
        return False


def build_device_rows_from_interfaces(interface_rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """Search API fallback에서는 device 상세가 없어 monitor_ip 기준 placeholder Device_List를 만든다."""
    rows: list[dict[str, object]] = []
    seen: set[str] = set()
    for interface in interface_rows:
        monitor_ip = str(interface.get("monitor_ip") or "")
        if not monitor_ip or monitor_ip in seen:
            continue
        seen.add(monitor_ip)
        rows.append(
            {
                "device_name": None,
                "device_ip": monitor_ip,
                "device_type": None,
                "status": "확인 필요",
            }
        )
    return rows


def build_interface_rows_from_targets(targets: list[SearchTarget]) -> list[dict[str, object]]:
    """Search API target을 Interface_List sheet row로 변환한다."""
    deduplicated = deduplicate_search_targets(targets)
    return [
        {
            "device_name": None,
            "device_ip": target.ip_address,
            "monitor_ip": target.ip_address,
            "interface_no": target.ifn,
            "interface_name": None,
            "interface_description": None,
            "speed": None,
            "interface_key": build_interface_key(target.ip_address, target.ifn),
        }
        for target in deduplicated
    ]


def fetch_topnconv_rows_for_interface(
    settings: NetscoutSettings,
    interface_row: dict[str, object],
    duration: str,
    top_n: int,
) -> list[dict[str, object]]:
    """한 monitor interface에서 TOPNCONV raw row를 가져온다."""
    monitor_ip = str(interface_row.get("deviceIPAddress") or interface_row.get("monitor_ip") or "")
    interface_no = str(interface_row.get("interfaceNumber") or interface_row.get("interface_no") or "")
    xml_text = build_topnconv_query_xml(monitor_ip, interface_no, duration, top_n)
    url = _url(settings.base_url, settings.data_api_endpoint)
    host = urlparse(url).hostname or "unknown-host"
    logger.info(
        "Calling registration TOPNCONV host=%s endpoint=%s monitor_ip=%s ifn=%s top_n=%s verify_ssl=%s",
        host,
        settings.data_api_endpoint,
        monitor_ip,
        interface_no,
        top_n,
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
            data=xml_text.encode("utf-8"),
            verify=settings.verify_ssl,
            timeout=settings.timeout,
        )
        elapsed = perf_counter() - start
    except requests.exceptions.Timeout as exc:
        raise RegistrationCandidatesError(f"TOPNCONV request timed out for monitor_ip={monitor_ip} ifn={interface_no}.") from exc
    except requests.exceptions.SSLError as exc:
        raise RegistrationCandidatesError(f"TOPNCONV SSL error for monitor_ip={monitor_ip} ifn={interface_no}.") from exc
    except requests.exceptions.ConnectionError as exc:
        raise RegistrationCandidatesError(f"TOPNCONV connection failed for monitor_ip={monitor_ip} ifn={interface_no}.") from exc
    except requests.exceptions.RequestException as exc:
        raise RegistrationCandidatesError(f"TOPNCONV request failed for monitor_ip={monitor_ip} ifn={interface_no}.") from exc

    logger.info("TOPNCONV response status=%s elapsed=%.3fs monitor_ip=%s ifn=%s", response.status_code, elapsed, monitor_ip, interface_no)
    if response.status_code != 200:
        raise RegistrationCandidatesError(
            f"TOPNCONV failed with HTTP {response.status_code}. Response preview={_sanitized_preview(response.text)!r}"
        )

    try:
        payload = response.json()
    except ValueError as exc:
        raise RegistrationCandidatesError(f"TOPNCONV returned invalid JSON for monitor_ip={monitor_ip} ifn={interface_no}.") from exc

    rows = payload.get("response", [])
    if not isinstance(rows, list):
        raise RegistrationCandidatesError("TOPNCONV JSON response field is not a list.")
    return [row for row in rows if isinstance(row, dict)]


def normalize_topnconv_raw_rows(
    rows: list[dict[str, object]],
    interface_row: dict[str, object],
    duration: str,
) -> list[dict[str, object]]:
    """TOPNCONV 응답에 interface metadata와 bps 계산값을 붙인다."""
    duration_seconds = duration_to_seconds(duration)
    normalized: list[dict[str, object]] = []
    for row in rows:
        octets = row.get("octets")
        normalized.append(
            {
                "deviceName": interface_row.get("deviceName") or interface_row.get("device_name"),
                "deviceIPAddress": interface_row.get("deviceIPAddress") or interface_row.get("device_ip"),
                "interfaceName": interface_row.get("interfaceName") or interface_row.get("interface_name"),
                "interfaceNumber": interface_row.get("interfaceNumber") or interface_row.get("interface_no"),
                "interface_key": interface_row.get("interface_key"),
                "appId": row.get("appId"),
                "appId_String": row.get("appId_String"),
                "srcAddress": row.get("srcAddress"),
                "srcPort": row.get("srcPort"),
                "destAddress": row.get("destAddress"),
                "destPort": row.get("destPort"),
                "srcToDestOctets": row.get("srcToDestOctets"),
                "destToSrcOctets": row.get("destToSrcOctets"),
                "octets": octets,
                "srcToDestPackets": row.get("srcToDestPackets"),
                "destToSrcPackets": row.get("destToSrcPackets"),
                "packets": row.get("packets"),
                "bps": calculate_bps(octets, duration_seconds),
            }
        )
    return normalized


def calculate_bps(octets: object, duration_seconds: int | None) -> float | None:
    """bitRate 계열 field가 없을 때 octets 기반으로 bps를 계산한다."""
    if not duration_seconds:
        return None
    try:
        return float(octets) * 8 / duration_seconds
    except (TypeError, ValueError):
        return None


def select_interface_top_n(rows: list[dict[str, object]], top_n: int) -> list[dict[str, object]]:
    """interface_key별 bps 내림차순 Top N row를 만든다."""
    grouped: dict[str, list[dict[str, object]]] = {}
    for row in rows:
        grouped.setdefault(str(row.get("interface_key") or ""), []).append(row)

    selected: list[dict[str, object]] = []
    for interface_key in sorted(grouped):
        sorted_rows = sorted(grouped[interface_key], key=lambda item: _sort_metric(item), reverse=True)
        for rank, row in enumerate(sorted_rows[:top_n], start=1):
            selected.append(
                {
                    "deviceName": row.get("deviceName"),
                    "deviceIPAddress": row.get("deviceIPAddress"),
                    "interfaceName": row.get("interfaceName"),
                    "interfaceNumber": row.get("interfaceNumber"),
                    "interface_key": row.get("interface_key"),
                    "rank": rank,
                    "appId_String": row.get("appId_String"),
                    "srcAddress": row.get("srcAddress"),
                    "srcPort": row.get("srcPort"),
                    "destAddress": row.get("destAddress"),
                    "destPort": row.get("destPort"),
                    "octets": row.get("octets"),
                    "packets": row.get("packets"),
                    "bps": row.get("bps"),
                }
            )
    return selected


def build_application_lookup(applications: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """application_type_code와 application_short_name 모두로 catalog lookup을 만든다."""
    lookup: dict[str, dict[str, Any]] = {}
    for application in applications:
        for field in ("application_type_code", "application_short_name"):
            value = application.get(field)
            if value not in (None, ""):
                lookup.setdefault(str(value).upper(), application)
    return lookup


def build_registered_ports(applications: list[dict[str, Any]]) -> dict[int, dict[str, Any]]:
    """Application catalog의 applicationPort/additionalPort/serverPortRange를 port lookup으로 만든다."""
    ports: dict[int, dict[str, Any]] = {}
    for application in applications:
        for field in ("application_port", "additional_port", "serverPortRange", "server_port_range"):
            for port in parse_port_values(application.get(field)):
                ports.setdefault(port, application)
    return ports


def parse_port_values(value: object) -> set[int]:
    """comma-separated port와 단순 range 문자열을 정수 port 집합으로 변환한다."""
    if value in (None, ""):
        return set()
    ports: set[int] = set()
    for part in str(value).replace(";", ",").split(","):
        token = part.strip()
        if not token:
            continue
        if "-" in token:
            start_text, end_text = token.split("-", 1)
            try:
                start = int(start_text.strip())
                end = int(end_text.strip())
            except ValueError:
                continue
            if 0 <= start <= end <= 65535 and end - start <= 4096:
                ports.update(range(start, end + 1))
            continue
        try:
            port = int(token)
        except ValueError:
            continue
        if 0 <= port <= 65535:
            ports.add(port)
    return ports


def filter_unregistered_application_rows(
    rows: list[dict[str, object]],
    application_lookup: dict[str, dict[str, Any]],
) -> list[dict[str, object]]:
    """이미 catalog에 등록된 Application을 제외하되 TCP_OTHER는 분석 대상으로 유지한다."""
    retained: list[dict[str, object]] = []
    for row in rows:
        app_name = str(row.get("appId_String") or "").upper()
        if app_name == "TCP_OTHER" or app_name not in application_lookup:
            retained.append(row)
    return retained


def annotate_candidate_ports(
    rows: list[dict[str, object]],
    registered_ports: dict[int, dict[str, Any]],
) -> list[dict[str, object]]:
    """TCP_OTHER row에 후보 port와 등록 여부 판단 field를 붙인다."""
    annotated: list[dict[str, object]] = []
    for row in rows:
        copied = dict(row)
        candidate_port, candidate_port_type = select_candidate_port(row.get("srcPort"), row.get("destPort"))
        copied["candidate_port"] = candidate_port
        copied["candidate_port_type"] = candidate_port_type
        if candidate_port is None:
            copied["port_registered"] = "unknown"
            copied["matched_application"] = None
            copied["registration_reason"] = "TCP_OTHER with unknown port"
        elif candidate_port in registered_ports:
            matched = registered_ports[candidate_port]
            copied["port_registered"] = True
            copied["matched_application"] = matched.get("application_short_name") or matched.get("application_type_code")
            copied["registration_reason"] = "TCP_OTHER but port is already registered"
        else:
            copied["port_registered"] = False
            copied["matched_application"] = None
            copied["registration_reason"] = f"TCP_OTHER with unregistered {candidate_port_type} port"
        annotated.append(copied)
    return annotated


def select_candidate_port(src_port: object, dest_port: object) -> tuple[int | None, str]:
    """destPort를 우선 사용하고, 0이면 srcPort를 후보 port로 사용한다."""
    dest = _to_int(dest_port)
    src = _to_int(src_port)
    if dest not in (None, 0):
        return dest, "destPort"
    if src not in (None, 0):
        return src, "srcPort"
    return None, "unknown"


def build_registration_candidates(rows: list[dict[str, object]]) -> list[dict[str, object]]:
    """TCP_OTHER이면서 미등록 port인 row만 최종 등록 후보로 변환한다."""
    candidates: list[dict[str, object]] = []
    for row in rows:
        if str(row.get("appId_String") or "").upper() != "TCP_OTHER":
            continue
        if row.get("candidate_port") in (None, "") or row.get("port_registered") is not False:
            continue
        candidate_port = row.get("candidate_port")
        candidates.append(
            {
                "deviceName": row.get("deviceName"),
                "deviceIPAddress": row.get("deviceIPAddress"),
                "interfaceName": row.get("interfaceName"),
                "interfaceNumber": row.get("interfaceNumber"),
                "interface_key": row.get("interface_key"),
                "candidate_application_name": f"TCP_OTHER_{candidate_port}",
                "candidate_port": candidate_port,
                "candidate_port_type": row.get("candidate_port_type"),
                "protocol": "TCP",
                "srcAddress": row.get("srcAddress"),
                "srcPort": row.get("srcPort"),
                "destAddress": row.get("destAddress"),
                "destPort": row.get("destPort"),
                "octets": row.get("octets"),
                "packets": row.get("packets"),
                "bps": row.get("bps"),
                "rank": row.get("rank"),
                "registration_reason": row.get("registration_reason"),
            }
        )
    return candidates


def build_summary_rows(
    devices: list[dict[str, object]],
    interfaces: list[dict[str, object]],
    topn_rows: list[dict[str, object]],
    top_rows: list[dict[str, object]],
    candidates: list[dict[str, object]],
    duration: str,
    top_n: int,
) -> list[dict[str, object]]:
    """Summary sheet용 metric/value row를 만든다."""
    tcp_other_rows = [row for row in top_rows if str(row.get("appId_String") or "").upper() == "TCP_OTHER"]
    top_ports = [str(row.get("candidate_port")) for row in candidates[:10] if row.get("candidate_port") not in (None, "")]
    return [
        {"metric": "total_devices", "value": len(devices)},
        {"metric": "total_interfaces", "value": len(interfaces)},
        {"metric": "total_topn_rows", "value": len(topn_rows)},
        {"metric": "tcp_other_rows", "value": len(tcp_other_rows)},
        {"metric": "registration_candidate_count", "value": len(candidates)},
        {"metric": "top_candidate_ports", "value": ", ".join(top_ports)},
        {"metric": "generated_at", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
        {"metric": "duration", "value": duration},
        {"metric": "top_n", "value": top_n},
    ]


def export_registration_candidates_xlsx(
    devices: list[dict[str, object]],
    interfaces: list[dict[str, object]],
    applications: list[dict[str, Any]],
    topn_rows: list[dict[str, object]],
    top_rows: list[dict[str, object]],
    candidates: list[dict[str, object]],
    summary: list[dict[str, object]],
    output_path: str | Path,
) -> Path:
    """Application 등록 후보 분석 workbook을 저장한다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    _write_sheet(workbook.active, "Device_List", DEVICE_LIST_COLUMNS, devices)
    _write_sheet(workbook.create_sheet("Interface_List"), "Interface_List", INTERFACE_LIST_COLUMNS, interfaces)
    _write_sheet(workbook.create_sheet("Application_List"), "Application_List", APPLICATION_COLUMNS, applications)
    _write_sheet(workbook.create_sheet("TOPNCONV_Raw"), "TOPNCONV_Raw", TOPNCONV_RAW_COLUMNS, topn_rows)
    _write_sheet(workbook.create_sheet("Interface_Top10_By_Bps"), "Interface_Top10_By_Bps", INTERFACE_TOPN_COLUMNS, top_rows)
    _write_sheet(workbook.create_sheet("Registration_Candidates"), "Registration_Candidates", REGISTRATION_CANDIDATE_COLUMNS, candidates)
    _write_sheet(workbook.create_sheet("Summary"), "Summary", SUMMARY_COLUMNS, summary)
    workbook.save(path)
    return path


def generate_registration_candidates_report(
    settings: NetscoutSettings,
    duration: str,
    top_n: int,
    output_path: str | Path,
    active_only: bool = True,
    max_devices: int | None = None,
    max_interfaces: int | None = None,
) -> RegistrationCandidatesResult:
    """Config API catalog와 Search API fallback interface 기반으로 등록 후보 report를 만든다."""
    try:
        client = NetscoutClient(settings)
        applications = parse_applications_response(client.get_applications_raw())
        devices = parse_devices_response(client.get_devices())
    except (NetscoutClientError, ParserError) as exc:
        raise RegistrationCandidatesError("Failed to collect Config API Application/Device List.") from exc

    try:
        interface_rows = collect_interfaces_for_devices(
            settings,
            devices,
            active_only=active_only,
            max_devices=max_devices,
            max_interfaces=max_interfaces,
        )
    except RegistrationCandidatesError as exc:
        raise RegistrationCandidatesError("Failed to collect Device Interface List.") from exc

    device_rows = [device for device in devices if not active_only or is_active_status(device.get("status"))]
    topn_rows: list[dict[str, object]] = []
    for interface_row in interface_rows:
        try:
            raw_rows = fetch_topnconv_rows_for_interface(settings, interface_row, duration, top_n)
            topn_rows.extend(normalize_topnconv_raw_rows(raw_rows, interface_row, duration))
        except RegistrationCandidatesError as exc:
            logger.warning(
                "Skipping failed TOPNCONV interface monitor_ip=%s ifn=%s: %s",
                interface_row.get("deviceIPAddress"),
                interface_row.get("interfaceNumber"),
                exc,
            )

    top_rows = select_interface_top_n(topn_rows, top_n)
    application_lookup = build_application_lookup(applications)
    retained_rows = filter_unregistered_application_rows(top_rows, application_lookup)
    annotated_rows = annotate_candidate_ports(retained_rows, build_registered_ports(applications))
    candidates = build_registration_candidates(annotated_rows)
    summary = build_summary_rows(device_rows, interface_rows, topn_rows, top_rows, candidates, duration, top_n)
    path = export_registration_candidates_xlsx(
        device_rows,
        interface_rows,
        applications,
        topn_rows,
        top_rows,
        candidates,
        summary,
        output_path,
    )
    return RegistrationCandidatesResult(
        devices_count=len(device_rows),
        interfaces_count=len(interface_rows),
        topn_rows_count=len(topn_rows),
        tcp_other_rows_count=len([row for row in top_rows if str(row.get("appId_String") or "").upper() == "TCP_OTHER"]),
        registration_candidate_count=len(candidates),
        output_path=path,
    )


def _write_sheet(worksheet, title: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
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


def _sort_metric(row: dict[str, object]) -> float:
    bps = row.get("bps")
    if bps not in (None, ""):
        return float(bps)
    value = row.get("octets")
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0


def _to_int(value: object) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _url(base_url: str, endpoint: str) -> str:
    return urljoin(base_url.rstrip("/") + "/", endpoint.lstrip("/"))


def _sanitized_preview(text: str, limit: int = 300) -> str:
    compact = " ".join((text or "").split())
    try:
        parsed = json.loads(compact)
    except json.JSONDecodeError:
        return compact[:limit]
    if isinstance(parsed, dict):
        for key in list(parsed):
            if key.lower() in {"password", "token", "cookie", "authorization", "apikey", "api_key"}:
                parsed[key] = "***"
    return json.dumps(parsed, ensure_ascii=True)[:limit]

