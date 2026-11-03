"""NETSCOUT Automation v0.86 핵심 로직.

v0.86은 V8.1/V8.5 계열의 Inventory/Top/Detail 기능은 유지하고, 사용자가 V55_TEST에서
확인한 APPLICATION raw 응답 구조를 기준으로 트래픽 분석을 수행한다.

주요 흐름은 세 가지다.

1. Inventory: Config API로 Application / Device / Interface 목록을 조회한다.
2. Top N: Device / Interface 기준 APPLICATION 트래픽을 조회해 bitRate 기준 Top N을 만든다.
3. Detail: Top N에서 고른 Application의 Source/Destination IP와 Port 상세를 조회한다.
4. Output: 실행 날짜 기준 `output/YYYYMMDD/<stage>` 구조로 결과를 저장한다.
"""

from __future__ import annotations

import json
import logging
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from time import perf_counter
from typing import Any
from urllib.parse import urlencode, urljoin, urlparse
from xml.sax.saxutils import escape

import requests
import urllib3
import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from requests.auth import HTTPBasicAuth

from netscout_automation.client import NetscoutClient
from netscout_automation.config import NetscoutSettings
from netscout_automation.parsers import (
    APPLICATION_COLUMNS,
    DEVICE_COLUMNS,
    INTERFACE_COLUMNS,
    parse_applications_response,
    parse_device_interfaces_response,
    parse_devices_response,
)

logger = logging.getLogger(__name__)

DEFAULT_MIN_BITRATE = 1_000_000
DEFAULT_TOP_N = 10
STEP2_BIT_RATE_COLUMN = "Bit Rate"
UNDEFINED_KEYWORDS = ("OTHER", "UNKNOWN")
DEFAULT_DRILLDOWN_APPS = ("HTTPS", "HTTP", "TCP_OTHER", "UDP_OTHER", "QUIC")

CANDIDATE_COLUMNS = [
    "Rank",
    "Application",
    "Application Type",
    "Application Type Code",
    "Parent Application",
    "Application Group",
    "Traffic",
    "Packets",
    "Source",
    "Defined",
    "Status",
    "Action",
    "Reason",
    "App ID",
    "Application Key",
    "Port",
]

APPLICATION_INVENTORY_COLUMNS = [
    "ApplicationShortName",
    "ApplicationLongName",
    "ApplicationType",
    "ApplicationTypeCode",
    "ParentApplication",
    "ApplicationPort",
    "AdditionalPort",
    "ApplicationGroup",
    "IsActive",
    "Logged",
    "ResponseTime",
    "Session",
    "RecordingType",
]

TRAFFIC_RAW_COLUMNS = [
    "monitor_ip",
    "ifn",
    "sessionId",
    "appId",
    "appId_String",
    "srcAddress",
    "srcPort",
    "destAddress",
    "destPort",
    "flowAddressType",
    "srcToDestOctets",
    "destToSrcOctets",
    "srcToDestPackets",
    "destToSrcPackets",
    "bitRate",
    "bitRateIn",
    "bitRateOut",
]

APPLICATION_SUMMARY_COLUMNS = [
    "rank",
    "Application",
    "App ID",
    "Total bitRate",
    "Total Mbps",
    "Total bitRate In",
    "Total bitRate Out",
    "Total Byte",
    "Total Byte In",
    "Total Byte Out",
    "Total Packets",
    "Total Packets In",
    "Total Packets Out",
    "Flow Count",
]

UNDEFINED_CANDIDATE_COLUMNS = [
    "rank",
    "Application",
    "App ID",
    "Total bitRate",
    "Total Mbps",
    "Total Byte",
    "Flow Count",
    "candidate_reason",
]

RUN_SUMMARY_COLUMNS = ["metric", "value"]
STEP2_COLUMNS = [
    "Rank",
    "Application Name",
    "Application Type Code",
    "Application Type",
    "Parent Application",
    "Application Group",
    "App ID",
    STEP2_BIT_RATE_COLUMN,
    "Bytes",
    "Packets",
    "Device Name",
    "Device IP",
    "Interface Name",
    "Interface Alias",
    "Interface Number",
]
STEP3_COLUMNS = [
    "Application Name",
    "App ID",
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "bitRate",
    "bitRateIn",
    "bitRateOut",
    "Bytes",
    "Packets",
    "Device Name",
    "Device IP",
    "Interface Name",
    "Interface Alias",
    "Interface Number",
]
UNIQUE_IP_PAIR_COLUMNS = [
    "Application Name",
    "Source IP",
    "Destination IP",
    "Max bitRate",
    "Total bitRate",
    "Flow Count",
    "Ports Observed",
    "Device Name",
    "Device IP",
    "Interface Name",
    "Interface Alias",
    "Interface Number",
]
UNIQUE_IP_PORT_PAIR_COLUMNS = [
    "Application Name",
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "Max bitRate",
    "Total bitRate",
    "Flow Count",
    "Device Name",
    "Device IP",
    "Interface Name",
    "Interface Alias",
    "Interface Number",
]

# STEP 2의 Bit Rate는 Application 단위 Data API 응답의 bitRate를 그대로 출력한다.
# STEP 2는 Python에서 Bit Rate를 합산하거나 Bytes 기반으로 환산하지 않는다.
# STEP 3의 bitRate/bitRateIn/bitRateOut은 개별 Source/Destination 통신 row의 값이다.
# Device IP / Interface Number는 조회 기준이 되는 관측 지점이고,
# Source IP / Destination IP는 그 관측 지점에서 보인 실제 conversation endpoint다.


class V81Error(RuntimeError):
    """V8 처리 중 실패했을 때 발생한다."""


@dataclass(frozen=True)
class Inventory:
    applications: list[dict[str, Any]]
    devices: list[dict[str, Any]]
    interfaces: list[dict[str, Any]]


@dataclass(frozen=True)
class TrafficResult:
    raw_rows_count: int
    app_summary_count: int
    candidate_count: int
    output_path: Path


@dataclass(frozen=True)
class StepResult:
    rows_count: int
    output_path: Path


@dataclass(frozen=True)
class TargetYamlResult:
    targets_count: int
    output_path: Path


@dataclass(frozen=True)
class BatchTopResult:
    planned_count: int
    success_count: int
    failure_count: int
    skipped_count: int
    results: list[dict[str, Any]]


@dataclass(frozen=True)
class DetailTargetsYamlResult:
    targets_count: int
    output_path: Path


@dataclass(frozen=True)
class BatchDetailResult:
    planned_count: int
    success_count: int
    failure_count: int
    skipped_count: int
    results: list[dict[str, Any]]


@dataclass(frozen=True)
class TimeDef:
    duration: str | None = None
    start_time: str | None = None
    end_time: str | None = None
    resolution: str = "NO_RESOLUTION"


def build_time_def(duration: str | None, start_time: str | None, end_time: str | None, resolution: str) -> TimeDef:
    """duration 방식 또는 start/end 절대시간 방식 중 하나의 TimeDef를 만든다."""
    if duration and (start_time or end_time):
        raise V81Error("duration과 start-time/end-time은 동시에 사용할 수 없다.")
    if bool(start_time) != bool(end_time):
        raise V81Error("start-time과 end-time은 함께 지정해야 한다.")
    if start_time and end_time:
        return TimeDef(start_time=start_time, end_time=end_time, resolution=resolution)
    return TimeDef(duration=duration or "LAST_HOUR", resolution=resolution)


def time_def_label(time_def: TimeDef) -> str:
    """파일명과 Summary에 사용할 시간 범위 표시값을 만든다."""
    if time_def.start_time and time_def.end_time:
        return f"{time_def.start_time}_to_{time_def.end_time}"
    return str(time_def.duration or "LAST_HOUR")


def config_get_collection(settings: NetscoutSettings, endpoint: str, action: str) -> list[dict[str, Any]]:
    """Config API GET 응답을 JSON/XML 모두 허용해 list[dict]로 변환한다."""
    raw = NetscoutClient(settings)._get_raw(endpoint, action)
    return parse_config_collection(raw)


def parse_config_collection(raw: str) -> list[dict[str, Any]]:
    """Discovery API가 XML 또는 JSON으로 반환한 collection을 list[dict]로 정규화한다."""
    text = raw.strip()
    if not text:
        return []
    if text.startswith("<"):
        return parse_xml_collection(text)
    try:
        payload = json.loads(text)
    except json.JSONDecodeError as exc:
        raise V81Error("Config API response is neither JSON nor XML.") from exc
    return collection_from_object(payload)


def parse_xml_collection(text: str) -> list[dict[str, Any]]:
    """단순 NETSCOUT XML tree를 row list로 변환한다."""
    try:
        root = ET.fromstring(text)
    except ET.ParseError as exc:
        raise V81Error("Invalid XML discovery response.") from exc
    children = list(root)
    if not children:
        return []
    if all(len(child) == 0 for child in children):
        return [xml_element_to_dict(root)]
    return [xml_element_to_dict(child) for child in children if isinstance(xml_element_to_dict(child), dict)]


def xml_element_to_dict(element: ET.Element) -> dict[str, Any] | str:
    children = list(element)
    if not children:
        return (element.text or "").strip()
    result: dict[str, Any] = {}
    for child in children:
        key = child.tag.rsplit("}", 1)[-1]
        value = xml_element_to_dict(child)
        if key in result:
            if not isinstance(result[key], list):
                result[key] = [result[key]]
            result[key].append(value)
        else:
            result[key] = value
    return result


def collection_from_object(payload: Any) -> list[dict[str, Any]]:
    """알려진 wrapper key가 달라도 첫 번째 list[dict]를 안전하게 찾는다."""
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if not isinstance(payload, dict):
        return []
    for value in payload.values():
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            nested = collection_from_object(value)
            if nested:
                return nested
    return [payload]


def collect_inventory(settings: NetscoutSettings) -> Inventory:
    """V5와 동일하게 Config API로 Application / Device / Interface 목록을 수집한다."""
    client = NetscoutClient(settings)
    # Config API 인증은 공통 NetscoutClient가 처리한다. 비밀번호는 settings에만 있고 로그로 출력하지 않는다.
    applications = parse_applications_response(client.get_applications_raw())
    devices = parse_devices_response(client.get_devices())
    interfaces: list[dict[str, Any]] = []
    for device in devices:
        device_name = str(device.get("deviceName") or "")
        if not device_name:
            continue
        try:
            parsed_interfaces = parse_device_interfaces_response(
                client.get_device_interfaces(device_name),
                device_name=device_name,
                device_ip_address=device.get("deviceIPAddress"),
            )
            interfaces.extend(parsed_interfaces)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Skipping failed Device Interface List device=%s: %s", device_name, exc)
    return Inventory(applications=applications, devices=devices, interfaces=interfaces)


def select_device_interface(
    inventory: Inventory,
    device_name: str,
    interface_filter: str | None,
) -> dict[str, Any]:
    """Device 이름과 Interface 이름/번호로 STEP 2/3 조회 대상을 하나 선택한다."""
    matched_device_names = resolve_device_names(inventory, device_name)
    matched = []
    for interface in inventory.interfaces:
        if str(interface.get("deviceName") or "") not in matched_device_names:
            continue
        if interface_filter and not interface_matches(interface, interface_filter):
            continue
        if not interface.get("deviceIPAddress") or interface.get("interfaceNumber") in (None, ""):
            continue
        matched.append(interface)
    if not matched:
        candidates = [
            f"{item.get('deviceName')} / {item.get('interfaceName')} / {item.get('interfaceNumber')}"
            for item in inventory.interfaces
            if str(item.get("deviceName") or "") in matched_device_names
        ]
        hint = "; candidates: " + ", ".join(candidates[:10]) if candidates else ""
        raise V81Error(f"No matching interface. Check device name and interface name/number.{hint}")
    return matched[0]


def resolve_device_names(inventory: Inventory, device_name: str) -> set[str]:
    """입력 Device 값을 실제 deviceName 목록으로 자동 매칭한다.

    우선순위는 exact -> alias/IP exact -> 특수문자 제거 exact -> 부분 포함이다.
    부분 포함 결과가 여러 개면 잘못된 장비 선택을 막기 위해 에러로 후보를 보여준다.
    """
    wanted = str(device_name or "").strip()
    wanted_lower = wanted.lower()
    wanted_key = compact_match_key(wanted)
    if not wanted:
        raise V81Error("Device name is required.")

    exact = [
        str(device.get("deviceName") or "")
        for device in inventory.devices
        if str(device.get("deviceName") or "").strip().lower() == wanted_lower
    ]
    if exact:
        return set(exact)

    alias_or_ip = [
        str(device.get("deviceName") or "")
        for device in inventory.devices
        if str(device.get("alias") or "").strip().lower() == wanted_lower
        or str(device.get("deviceIPAddress") or "").strip().lower() == wanted_lower
    ]
    if alias_or_ip:
        return set(alias_or_ip)

    compact = [
        str(device.get("deviceName") or "")
        for device in inventory.devices
        if compact_match_key(device.get("deviceName")) == wanted_key
        or compact_match_key(device.get("alias")) == wanted_key
    ]
    if compact:
        return set(compact)

    partial = [
        str(device.get("deviceName") or "")
        for device in inventory.devices
        if wanted_key
        and (
            wanted_key in compact_match_key(device.get("deviceName"))
            or compact_match_key(device.get("deviceName")) in wanted_key
            or wanted_key in compact_match_key(device.get("alias"))
            or compact_match_key(device.get("alias")) in wanted_key
        )
    ]
    partial = list(dict.fromkeys(name for name in partial if name))
    if len(partial) == 1:
        return set(partial)
    if len(partial) > 1:
        raise V81Error(f"Ambiguous device name '{device_name}'. Candidates: {', '.join(partial[:10])}")

    return {wanted}


def compact_match_key(value: object) -> str:
    """사람이 입력한 장비명 비교용으로 영문/숫자만 남긴다."""
    return "".join(ch for ch in str(value or "").lower() if ch.isalnum())


def interface_matches(interface: dict[str, Any], filter_text: str) -> bool:
    """Interface 이름, 번호, if번호 형식 중 하나가 입력값과 일치하는지 확인한다."""
    wanted = str(filter_text or "").strip().lower()
    number = str(interface.get("interfaceNumber") or "").strip().lower()
    candidates = {
        str(interface.get("interfaceName") or "").strip().lower(),
        number,
        f"if{number}" if number else "",
    }
    return wanted in candidates


def find_application_id(applications: list[dict[str, Any]], application_name: str) -> str:
    """Application 이름에 해당하는 appId 필터값을 찾고, 없으면 입력값을 그대로 사용한다."""
    wanted = str(application_name or "").strip().upper()
    for app in applications:
        # 사람이 입력하는 값은 short name, long name, type code가 섞일 수 있어서 세 필드를 모두 비교한다.
        for field in ("application_type_code", "application_short_name", "application_long_name"):
            if str(app.get(field) or "").strip().upper() == wanted:
                return str(app.get("application_type_code") or app.get("application_short_name") or application_name)
    return str(application_name)


def build_application_traffic_query_xml(
    monitor_ip: str,
    ifn: str,
    duration: str = "LAST_HOUR",
    use_topn: bool = False,
    top_n: int = 100,
    top_column: str = "bitRate",
    application: str | None = None,
) -> str:
    """V55_TEST의 확인된 XML을 기준으로 APPLICATION 트래픽 조회 XML을 만든다."""
    time_def = build_time_def(duration, None, None, "NO_RESOLUTION")
    return build_application_traffic_query_xml_for_time_def(
        monitor_ip,
        ifn,
        time_def,
        use_topn=use_topn,
        top_n=top_n,
        top_column=top_column,
        application=application,
    )


def build_application_traffic_query_xml_for_time_def(
    monitor_ip: str,
    ifn: str,
    time_def: TimeDef,
    use_topn: bool = False,
    top_n: int = 100,
    top_column: str = "bitRate",
    application: str | None = None,
) -> str:
    """duration 또는 start/end 방식 TimeDef를 사용해 APPLICATION 트래픽 조회 XML을 만든다.

    NetworkParameter=APPLICATION은 V55_TEST에서 실제 응답을 확인한 쿼리다.
    이 쿼리에서 appId_String, bitRate, src/dest IP, src/dest Port가 내려오므로
    STEP 2와 STEP 3 모두 같은 기본 XML 구조를 사용한다.
    """
    app_filter_xml = f"\n        <appId>{escape(str(application))}</appId>\n      " if application else ""
    time_xml = (
        f"<startTime>{escape(str(time_def.start_time))}</startTime>\n"
        f"    <endTime>{escape(str(time_def.end_time))}</endTime>"
        if time_def.start_time and time_def.end_time
        else f"<duration>{escape(str(time_def.duration))}</duration>"
    )
    function_xml = (
        f"""<FunctionList>
    <Function>
      <name>TopN</name>
      <column>{escape(str(top_column))}</column>
      <order>descending</order>
      <nValue>{int(top_n)}</nValue>
    </Function>
  </FunctionList>"""
        if use_topn
        else "<FunctionList />"
    )
    # appId 필터는 서버에 전달하지만, 응답에 다른 Application이 섞일 수 있어
    # STEP 3에서는 Python 쪽에서 appId_String/appId 기준으로 한 번 더 필터링한다.
    return f"""<GenericClientQuery>
  <NetworkObjectData>
    <NetworkParameter>APPLICATION</NetworkParameter>
    <SelectColumnList>
      <ClientColumn>sessionId</ClientColumn>
      <ClientColumn>appId</ClientColumn>
      <ClientColumn>appId_String</ClientColumn>
      <ClientColumn>srcAddress</ClientColumn>
      <ClientColumn>srcPort</ClientColumn>
      <ClientColumn>destAddress</ClientColumn>
      <ClientColumn>destPort</ClientColumn>
      <ClientColumn>flowAddressType</ClientColumn>
      <ClientColumn>srcToDestOctets</ClientColumn>
      <ClientColumn>destToSrcOctets</ClientColumn>
      <ClientColumn>srcToDestPackets</ClientColumn>
      <ClientColumn>destToSrcPackets</ClientColumn>
      <ClientColumn>bitRate</ClientColumn>
      <ClientColumn>bitRateIn</ClientColumn>
      <ClientColumn>bitRateOut</ClientColumn>
    </SelectColumnList>
  </NetworkObjectData>
  <FlowFilterList>
    <FlowFilter>
      <IPAddress>{escape(str(monitor_ip))}</IPAddress>
      <Ifn>{escape(str(ifn))}</Ifn>
      <FilterList>{app_filter_xml}</FilterList>
    </FlowFilter>
  </FlowFilterList>
  {function_xml}
  <TimeDef>
    {time_xml}
    <resolution>{escape(str(time_def.resolution))}</resolution>
  </TimeDef>
  <params>
    <columnOrder>true</columnOrder>
    <stringConversion>true</stringConversion>
  </params>
</GenericClientQuery>"""


def fetch_application_traffic_rows(
    settings: NetscoutSettings,
    monitor_ip: str,
    ifn: str,
    duration: str,
    use_topn: bool = False,
    top_n: int = 100,
    application: str | None = None,
) -> list[dict[str, Any]]:
    """Data API를 호출해 V8 APPLICATION 트래픽 raw row를 가져온다."""
    xml_text = build_application_traffic_query_xml(monitor_ip, ifn, duration, use_topn, top_n, application=application)
    return fetch_application_traffic_rows_with_xml(settings, monitor_ip, ifn, duration, use_topn, xml_text, application)


def fetch_application_traffic_rows_for_time_def(
    settings: NetscoutSettings,
    monitor_ip: str,
    ifn: str,
    time_def: TimeDef,
    use_topn: bool = False,
    top_n: int = 100,
    application: str | None = None,
) -> list[dict[str, Any]]:
    """TimeDef 객체를 사용해 Data API APPLICATION 트래픽 raw row를 가져온다."""
    xml_text = build_application_traffic_query_xml_for_time_def(monitor_ip, ifn, time_def, use_topn, top_n, application=application)
    return fetch_application_traffic_rows_with_xml(settings, monitor_ip, ifn, time_def_label(time_def), use_topn, xml_text, application)


def fetch_application_traffic_rows_with_xml(
    settings: NetscoutSettings,
    monitor_ip: str,
    ifn: str,
    duration_label: str,
    use_topn: bool,
    xml_text: str,
    application: str | None = None,
) -> list[dict[str, Any]]:
    """생성된 XML을 Data API로 전송한다.

    로그에는 host, endpoint, monitor_ip, ifn 같은 운영상 필요한 값만 남긴다.
    Basic Auth의 password/token/cookie 값은 절대 출력하지 않는다.
    """
    url = urljoin(settings.base_url.rstrip("/") + "/", settings.data_api_endpoint.lstrip("/"))
    host = urlparse(url).hostname or "unknown-host"
    logger.info(
        "Calling V8 APPLICATION traffic host=%s endpoint=%s monitor_ip=%s ifn=%s duration=%s application=%s use_topn=%s verify_ssl=%s",
        host,
        settings.data_api_endpoint,
        monitor_ip,
        ifn,
        duration_label,
        application or "",
        use_topn,
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
        raise V81Error("V8 APPLICATION traffic request timed out") from exc
    except requests.exceptions.SSLError as exc:
        raise V81Error("V8 APPLICATION traffic SSL error") from exc
    except requests.exceptions.ConnectionError as exc:
        raise V81Error("V8 APPLICATION traffic connection error") from exc
    except requests.exceptions.RequestException as exc:
        raise V81Error("V8 APPLICATION traffic request failed") from exc
    logger.info("V8 APPLICATION traffic status=%s elapsed=%.3fs", response.status_code, elapsed)
    if response.status_code == 204:
        return []
    if response.status_code != 200:
        raise V81Error(f"V8 APPLICATION traffic HTTP {response.status_code}. preview={response.text[:200]!r}")
    return parse_response_rows(response.json())


def load_raw_json(path: str | Path) -> list[dict[str, Any]]:
    """curl 등으로 저장한 raw JSON 파일에서 response row를 읽는다."""
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    return parse_response_rows(payload)


def parse_response_rows(payload: Any) -> list[dict[str, Any]]:
    """NETSCOUT JSON payload의 response 배열을 검증해 반환한다."""
    if not isinstance(payload, dict):
        raise V81Error("V8 response is not a JSON object")
    rows = payload.get("response", [])
    if not isinstance(rows, list):
        raise V81Error("V8 response field is not a list")
    return [row for row in rows if isinstance(row, dict)]


def normalize_traffic_rows(rows: list[dict[str, Any]], monitor_ip: str = "", ifn: str = "") -> list[dict[str, Any]]:
    """확인된 raw field를 Excel용 내부 구조로 정리한다.

    monitor_ip/ifn은 Device Interface 쪽 관측 지점 정보다.
    row 내부의 srcAddress/destAddress와는 의미가 다르므로 별도 필드로 유지한다.
    """
    normalized = []
    for row in rows:
        normalized.append({column: row.get(column) for column in TRAFFIC_RAW_COLUMNS} | {"monitor_ip": monitor_ip, "ifn": ifn})
    return normalized


def filter_application_rows(rows: list[dict[str, Any]], application: str | None = None, limit: int | None = None) -> list[dict[str, Any]]:
    """Application 이름 또는 appId 숫자로 raw row를 필터링하고 bitRate 내림차순으로 제한한다.

    Data API 서버사이드 appId 필터가 항상 원하는 row만 돌려준다고 가정하지 않는다.
    그래서 STEP 3에서는 여기서 appId_String/appId를 기준으로 다시 걸러낸다.
    """
    filtered = rows
    if application:
        wanted = str(application).strip().upper()
        filtered = [
            row
            for row in filtered
            if str(row.get("appId_String") or "").strip().upper() == wanted
            or str(row.get("appId") or "").strip().upper() == wanted
        ]
    filtered = sorted(filtered, key=lambda item: to_float(item.get("bitRate")) or 0, reverse=True)
    if limit is not None and limit > 0:
        return filtered[:limit]
    return filtered


def dedup_ip_pairs(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Source/Destination IP 조합 기준으로 중복을 제거하고 가장 높은 bitRate row를 유지한다.

    Source IP와 Destination IP가 둘 다 같은 row만 중복으로 본다.
    A -> B와 B -> A처럼 방향이 반대인 row는 서로 다른 통신으로 유지한다.
    """
    deduped: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    sorted_rows = sorted(rows, key=lambda item: to_float(item.get("bitRate")) or 0, reverse=True)
    for row in sorted_rows:
        src_ip = str(row.get("srcAddress") or "").strip()
        dest_ip = str(row.get("destAddress") or "").strip()
        key = (src_ip, dest_ip)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)
    return deduped




def build_step3_rows(rows: list[dict[str, Any]], interface: dict[str, Any]) -> list[dict[str, Any]]:
    """Application 상세 raw row를 STEP 3 요구 컬럼으로 변환한다.

    Source/Destination IP와 Port는 실제 conversation endpoint다.
    Device/Interface 컬럼은 이 conversation이 관측된 지점을 나타낸다.
    """
    result = []
    for row in rows:
        bytes_value = sum_number(row.get("srcToDestOctets"), row.get("destToSrcOctets"))
        packets_value = sum_number(row.get("srcToDestPackets"), row.get("destToSrcPackets"))
        result.append(
            {
                "Application Name": row.get("appId_String"),
                "App ID": row.get("appId"),
                "Source IP": row.get("srcAddress"),
                "Source Port": row.get("srcPort"),
                "Destination IP": row.get("destAddress"),
                "Destination Port": row.get("destPort"),
                "bitRate": row.get("bitRate"),
                "bitRateIn": row.get("bitRateIn"),
                "bitRateOut": row.get("bitRateOut"),
                "Bytes": bytes_value,
                "Packets": packets_value,
                "Device Name": interface.get("deviceName"),
                "Device IP": interface.get("deviceIPAddress"),
                "Interface Name": interface.get("interfaceName"),
                "Interface Alias": interface.get("alias"),
                "Interface Number": interface.get("interfaceNumber"),
            }
        )
    return sorted(result, key=lambda item: to_float(item.get("bitRate")) or 0, reverse=True)


def build_unique_ip_pair_rows(rows: list[dict[str, Any]], interface: dict[str, Any]) -> list[dict[str, Any]]:
    """Source/Destination IP pair 기준 요약 시트를 만든다.

    Detail 원본 row는 삭제하지 않고, eye checking용 추가 시트로만 사용한다.
    Port가 다른 같은 IP pair는 한 줄로 묶고, 관측된 port 조합은 Ports Observed에 모아둔다.
    """
    grouped: dict[tuple[str, str, str], dict[str, Any]] = {}
    for row in rows:
        application = str(row.get("appId_String") or "")
        src_ip = str(row.get("srcAddress") or "")
        dest_ip = str(row.get("destAddress") or "")
        key = (application, src_ip, dest_ip)
        bit_rate = to_float(row.get("bitRate")) or 0
        src_port = str(row.get("srcPort") if row.get("srcPort") is not None else "")
        dest_port = str(row.get("destPort") if row.get("destPort") is not None else "")
        item = grouped.setdefault(
            key,
            {
                "Application Name": application,
                "Source IP": src_ip,
                "Destination IP": dest_ip,
                "Max bitRate": 0,
                "Total bitRate": 0,
                "Flow Count": 0,
                "_ports": set(),
                "Device Name": interface.get("deviceName"),
                "Device IP": interface.get("deviceIPAddress"),
                "Interface Name": interface.get("interfaceName"),
                "Interface Alias": interface.get("alias"),
                "Interface Number": interface.get("interfaceNumber"),
            },
        )
        item["Max bitRate"] = max(to_float(item.get("Max bitRate")) or 0, bit_rate)
        item["Total bitRate"] = sum_number(item.get("Total bitRate"), bit_rate)
        item["Flow Count"] = int(item.get("Flow Count") or 0) + 1
        item["_ports"].add(f"{src_port}->{dest_port}")
    result = []
    for item in grouped.values():
        ports = sorted(str(port) for port in item.pop("_ports", set()) if str(port) != "->")
        item["Ports Observed"] = ", ".join(ports)
        result.append(item)
    return sorted(result, key=lambda item: to_float(item.get("Max bitRate")) or 0, reverse=True)


def build_unique_ip_port_pair_rows(rows: list[dict[str, Any]], interface: dict[str, Any]) -> list[dict[str, Any]]:
    """Source/Destination IP + Port pair 기준 요약 시트를 만든다.

    등록 후보 확인에서는 Port가 중요할 수 있으므로 IP pair 요약과 별도로 제공한다.
    """
    grouped: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    for row in rows:
        application = str(row.get("appId_String") or "")
        src_ip = str(row.get("srcAddress") or "")
        src_port = row.get("srcPort")
        dest_ip = str(row.get("destAddress") or "")
        dest_port = row.get("destPort")
        key = (application, src_ip, str(src_port), dest_ip, str(dest_port))
        bit_rate = to_float(row.get("bitRate")) or 0
        item = grouped.setdefault(
            key,
            {
                "Application Name": application,
                "Source IP": src_ip,
                "Source Port": src_port,
                "Destination IP": dest_ip,
                "Destination Port": dest_port,
                "Max bitRate": 0,
                "Total bitRate": 0,
                "Flow Count": 0,
                "Device Name": interface.get("deviceName"),
                "Device IP": interface.get("deviceIPAddress"),
                "Interface Name": interface.get("interfaceName"),
                "Interface Alias": interface.get("alias"),
                "Interface Number": interface.get("interfaceNumber"),
            },
        )
        item["Max bitRate"] = max(to_float(item.get("Max bitRate")) or 0, bit_rate)
        item["Total bitRate"] = sum_number(item.get("Total bitRate"), bit_rate)
        item["Flow Count"] = int(item.get("Flow Count") or 0) + 1
    return sorted(grouped.values(), key=lambda item: to_float(item.get("Max bitRate")) or 0, reverse=True)


def run_step1(settings: NetscoutSettings, output_dir: str | Path = ".", output_file: str | None = None) -> StepResult:
    """STEP 1: NETSCOUT Application / Device / Interface 현황을 Excel로 저장한다."""
    inventory = collect_inventory(settings)
    application_rows = format_application_inventory_rows(inventory.applications)
    output_path = staged_output_path(output_dir, "00_inventory", output_file, "Netscout_Inventory")
    path = write_xlsx(
        output_path,
        [
            ("Applications", APPLICATION_INVENTORY_COLUMNS, application_rows),
            ("Devices", DEVICE_COLUMNS, inventory.devices),
            ("Interfaces", INTERFACE_COLUMNS, inventory.interfaces),
        ],
    )
    append_run_summary(output_dir, f"inventory output={path} rows={len(inventory.applications) + len(inventory.devices) + len(inventory.interfaces)}")
    return StepResult(len(inventory.applications) + len(inventory.devices) + len(inventory.interfaces), path)


def build_targets_yaml_data(
    inventory: Inventory,
    duration: str,
    top_n: int,
    run_detail: bool = False,
    detail_top_n: int = 10,
    device_filter: str | None = None,
    interface_filter: str | None = None,
) -> dict[str, Any]:
    """Inventory 결과를 사람이 편집하기 쉬운 실행 대상 YAML 구조로 변환한다.

    이 YAML은 Top N 실행 대상을 고르는 1차 체크리스트다.
    Device/Interface가 여러 개면 모두 targets에 넣고, 사람은 enabled 값을 바꿔 실행 대상을 줄인다.
    """
    targets = []
    for interface in inventory.interfaces:
        device_name = str(interface.get("deviceName") or "")
        interface_name = str(interface.get("interfaceName") or "")
        interface_number = str(interface.get("interfaceNumber") or "")
        if device_filter and device_filter.lower() not in device_name.lower():
            continue
        if interface_filter and not interface_matches(interface, interface_filter):
            continue
        targets.append(
            {
                "enabled": True,
                "device": device_name,
                "device_ip": interface.get("deviceIPAddress"),
                "interface": interface_name or (f"if{interface_number}" if interface_number else ""),
                "interface_alias": interface.get("alias"),
                "interface_number": interface.get("interfaceNumber"),
                "ifn": interface.get("ifn") or interface.get("interfaceNumber"),
                "status": interface.get("status"),
                "duration": duration,
                "top_n": top_n,
                "run_detail": run_detail,
                "detail_top_n": detail_top_n,
                "include_applications": [],
                "exclude_applications": [],
                "memo": "",
            }
        )
    return {
        "version": 1,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "description": "Edit enabled/include_applications/exclude_applications before running batch.",
        "defaults": {
            "duration": duration,
            "top_n": top_n,
            "run_detail": run_detail,
            "detail_top_n": detail_top_n,
        },
        "targets": targets,
    }


def write_targets_yaml(data: dict[str, Any], output_path: str | Path) -> Path:
    """대상 YAML 파일을 저장한다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    return path


def run_targets_yaml(
    settings: NetscoutSettings,
    output_path: str | Path,
    duration: str,
    top_n: int,
    run_detail: bool = False,
    detail_top_n: int = 10,
    device_filter: str | None = None,
    interface_filter: str | None = None,
) -> TargetYamlResult:
    """Config API Inventory를 읽어 사람이 편집할 Batch target YAML을 생성한다."""
    inventory = collect_inventory(settings)
    data = build_targets_yaml_data(
        inventory,
        duration=duration,
        top_n=top_n,
        run_detail=run_detail,
        detail_top_n=detail_top_n,
        device_filter=device_filter,
        interface_filter=interface_filter,
    )
    path = write_targets_yaml(data, output_path)
    return TargetYamlResult(targets_count=len(data["targets"]), output_path=path)


def load_targets_yaml(path: str | Path) -> dict[str, Any]:
    """사람이 편집한 targets YAML 파일을 읽는다."""
    yaml_path = Path(path)
    if not yaml_path.exists():
        raise V81Error(f"Targets YAML file not found: {yaml_path}")
    data = yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise V81Error("Targets YAML root must be a mapping.")
    targets = data.get("targets")
    if not isinstance(targets, list):
        raise V81Error("Targets YAML must contain a targets list.")
    return data


def enabled_targets_from_yaml(data: dict[str, Any]) -> list[dict[str, Any]]:
    """targets YAML에서 enabled가 true인 항목만 반환한다."""
    targets = []
    for index, target in enumerate(data.get("targets", []), start=1):
        if not isinstance(target, dict):
            continue
        if target.get("enabled") is False:
            continue
        if not target.get("device"):
            raise V81Error(f"Target #{index} is missing device.")
        if not (target.get("interface") or target.get("interface_number") or target.get("ifn")):
            raise V81Error(f"Target #{index} is missing interface/interface_number/ifn.")
        targets.append(target)
    return targets


def target_interface_filter(target: dict[str, Any]) -> str:
    """Top 조회에 사용할 interface 입력값을 YAML target에서 고른다."""
    for key in ("interface", "interface_number", "ifn"):
        value = target.get(key)
        if value not in (None, ""):
            return str(value)
    return ""


def run_top_from_targets_yaml(
    settings: NetscoutSettings,
    targets_yaml: str | Path,
    output_dir: str | Path,
    dry_run: bool = False,
) -> BatchTopResult:
    """편집된 targets YAML을 읽어 enabled 대상마다 Top N을 실행한다.

    한 대상이 실패해도 전체 Batch를 중단하지 않는다.
    운영 중 일부 Interface만 실패할 수 있으므로 실패 건은 results/error에 남기고 다음 대상으로 진행한다.
    """
    data = load_targets_yaml(targets_yaml)
    targets = enabled_targets_from_yaml(data)
    results: list[dict[str, Any]] = []
    success_count = 0
    failure_count = 0
    for target in targets:
        device = str(target.get("device"))
        interface_filter = target_interface_filter(target)
        duration = str(target.get("duration") or data.get("defaults", {}).get("duration") or "LAST_7_DAYS")
        top_n = int(target.get("top_n") or data.get("defaults", {}).get("top_n") or 50)
        result_row = {
            "device": device,
            "interface": interface_filter,
            "duration": duration,
            "top_n": top_n,
            "dry_run": dry_run,
            "status": "planned" if dry_run else "",
            "output": "",
            "error": "",
        }
        if dry_run:
            results.append(result_row)
            continue
        try:
            result = run_step2(settings, device, interface_filter, duration, top_n, output_dir)
            result_row["status"] = "success"
            result_row["output"] = str(result.output_path)
            success_count += 1
        except Exception as exc:  # noqa: BLE001
            result_row["status"] = "failed"
            result_row["error"] = str(exc)
            failure_count += 1
            logger.warning("Skipping failed YAML target device=%s interface=%s: %s", device, interface_filter, exc)
        results.append(result_row)
    skipped_count = len(data.get("targets", [])) - len(targets)
    return BatchTopResult(
        planned_count=len(targets),
        success_count=success_count,
        failure_count=failure_count,
        skipped_count=skipped_count,
        results=results,
    )


def read_top_xlsx_rows(top_xlsx: str | Path) -> list[dict[str, Any]]:
    """Application_BPS_TopN 시트에서 row를 dict로 읽는다.

    Detail 대상 YAML은 사람이 이미 확인한 STEP 2 Excel을 입력으로 삼는다.
    그래서 여기서는 API를 호출하지 않고 Excel의 표시값만 읽는다.
    """
    path = Path(top_xlsx)
    if not path.exists():
        raise V81Error(f"Top Excel file not found: {path}")
    workbook = load_workbook(path, data_only=True)
    if "Application_BPS_TopN" not in workbook.sheetnames:
        raise V81Error("Top Excel must contain Application_BPS_TopN sheet.")
    worksheet = workbook["Application_BPS_TopN"]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(value) if value is not None else "" for value in rows[0]]
    result = []
    for values in rows[1:]:
        row = {headers[index]: values[index] if index < len(values) else None for index in range(len(headers))}
        if row.get("Application Name") in (None, ""):
            continue
        result.append(row)
    return result


def build_detail_targets_yaml_data(
    top_rows: list[dict[str, Any]],
    source_top_xlsx: str | Path,
    detail_top_n: int = 50,
    limit: int = 1000,
    dedup_ip: bool = False,
) -> dict[str, Any]:
    """TopN Excel row를 Detail 실행용 YAML 구조로 변환한다.

    Rank 1부터 detail_top_n까지 순서를 유지한다.
    bps/bytes/packets는 사람이 YAML을 볼 때 우선순위를 판단하는 참고값이며,
    Detail Data API 필터에는 application/device/interface/duration만 사용한다.
    """
    sorted_rows = sorted(top_rows, key=lambda row: to_float(row.get("Rank")) or 999999)
    selected_rows = sorted_rows[:detail_top_n] if detail_top_n > 0 else sorted_rows
    targets = []
    for row in selected_rows:
        targets.append(
            {
                "enabled": True,
                "rank": row.get("Rank"),
                "application": row.get("Application Type Code") or row.get("Application Name"),
                "app_id": row.get("App ID"),
                "bps": row.get("Bit Rate") or row.get("BPS"),
                "bytes": row.get("Bytes"),
                "packets": row.get("Packets"),
                "application_type": row.get("Application Type"),
                "application_type_code": row.get("Application Type Code"),
                "parent_application": row.get("Parent Application"),
                "application_group": row.get("Application Group"),
                "device": row.get("Device Name"),
                "device_ip": row.get("Device IP"),
                "interface": row.get("Interface Name"),
                "interface_alias": row.get("Interface Alias"),
                "interface_number": row.get("Interface Number"),
                "duration": "",
                "limit": limit,
                "dedup_ip": dedup_ip,
                "memo": "",
            }
        )
    return {
        "version": 1,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_top_xlsx": str(source_top_xlsx),
        "description": "Rank/Bit Rate are copied from Application_BPS_TopN for human review before detail batch.",
        "defaults": {
            "limit": limit,
            "dedup_ip": dedup_ip,
        },
        "targets": targets,
    }


def run_detail_targets_yaml(
    top_xlsx: str | Path,
    output_path: str | Path,
    detail_top_n: int = 50,
    limit: int = 1000,
    dedup_ip: bool = False,
) -> DetailTargetsYamlResult:
    """TopN Excel을 읽어 Detail 실행 대상 YAML을 만든다."""
    rows = read_top_xlsx_rows(top_xlsx)
    data = build_detail_targets_yaml_data(rows, top_xlsx, detail_top_n, limit, dedup_ip)
    path = write_targets_yaml(data, output_path)
    return DetailTargetsYamlResult(targets_count=len(data["targets"]), output_path=path)


def enabled_detail_targets_from_yaml(data: dict[str, Any]) -> list[dict[str, Any]]:
    """detail targets YAML에서 enabled가 true인 항목만 반환한다."""
    targets = []
    for index, target in enumerate(data.get("targets", []), start=1):
        if not isinstance(target, dict):
            continue
        if target.get("enabled") is False:
            continue
        for key in ("device", "application"):
            if not target.get(key):
                raise V81Error(f"Detail target #{index} is missing {key}.")
        if not (target.get("interface") or target.get("interface_number") or target.get("ifn")):
            raise V81Error(f"Detail target #{index} is missing interface/interface_number/ifn.")
        targets.append(target)
    return targets


def run_detail_from_targets_yaml(
    settings: NetscoutSettings,
    targets_yaml: str | Path,
    output_dir: str | Path,
    default_duration: str = "LAST_7_DAYS",
    dry_run: bool = False,
) -> BatchDetailResult:
    """편집된 detail targets YAML을 읽어 enabled 대상마다 Detail을 실행한다.

    Top N에서 고른 Application들을 사람이 검토한 뒤 한 번에 상세 조회하기 위한 Batch 실행이다.
    target별 duration이 비어 있으면 default_duration을 사용한다.
    """
    data = load_targets_yaml(targets_yaml)
    targets = enabled_detail_targets_from_yaml(data)
    results: list[dict[str, Any]] = []
    success_count = 0
    failure_count = 0
    for target in targets:
        device = str(target.get("device"))
        interface_filter = target_interface_filter(target)
        application = str(target.get("application"))
        duration = str(target.get("duration") or data.get("defaults", {}).get("duration") or default_duration)
        limit = int(target.get("limit") or data.get("defaults", {}).get("limit") or 1000)
        dedup_ip = bool(target.get("dedup_ip") if target.get("dedup_ip") is not None else data.get("defaults", {}).get("dedup_ip", False))
        result_row = {
            "rank": target.get("rank"),
            "application": application,
            "bps": target.get("bps"),
            "device": device,
            "interface": interface_filter,
            "duration": duration,
            "limit": limit,
            "dedup_ip": dedup_ip,
            "dry_run": dry_run,
            "status": "planned" if dry_run else "",
            "output": "",
            "error": "",
        }
        if dry_run:
            results.append(result_row)
            continue
        try:
            result = run_step3(settings, device, interface_filter, application, duration, output_dir, limit=limit, dedup_ip=dedup_ip)
            result_row["status"] = "success"
            result_row["output"] = str(result.output_path)
            success_count += 1
        except Exception as exc:  # noqa: BLE001
            result_row["status"] = "failed"
            result_row["error"] = str(exc)
            failure_count += 1
            logger.warning("Skipping failed detail YAML target device=%s interface=%s application=%s: %s", device, interface_filter, application, exc)
        results.append(result_row)
    skipped_count = len(data.get("targets", [])) - len(targets)
    return BatchDetailResult(
        planned_count=len(targets),
        success_count=success_count,
        failure_count=failure_count,
        skipped_count=skipped_count,
        results=results,
    )


def build_application_summary_query_xml(
    monitor_ip: str, ifn: str, time_def: TimeDef, top_n: int,
) -> str:
    """Application keys only; NETSCOUT returns its own bitRate and TopN.

    ASI v6.3.5: Creating Queries / Select Column List and KTI columns.
    Do not select conversation keys here: they change the response granularity.
    """
    if top_n <= 0:
        raise V81Error("top_n must be greater than zero")
    root = ET.fromstring(build_application_traffic_query_xml_for_time_def(
        monitor_ip, ifn, time_def, use_topn=True, top_n=top_n
    ))
    columns = root.find("./NetworkObjectData/SelectColumnList")
    columns.clear()
    for name in ("appId", "appId_String", "bitRate", "bitRateIn", "bitRateOut",
                 "octets", "packets"):
        ET.SubElement(columns, "ClientColumn").text = name
    return ET.tostring(root, encoding="unicode")


def build_step2_api_rows(
    raw_rows: list[dict[str, Any]], interface: dict[str, Any],
    applications: list[dict[str, Any]] | None = None,
) -> list[dict[str, Any]]:
    """Copy API values without summation, rate conversion, rounding or fallback."""
    lookup = application_lookup(applications or [])
    seen = set()
    result = []
    for rank, row in enumerate(raw_rows, start=1):
        key = str(row.get("appId") if row.get("appId") is not None
                  else row.get("appId_String"))
        if key in seen:
            raise V81Error("Application summary API returned duplicate Application rows; "
                           "check query granularity. Bit Rate will not be summed.")
        seen.add(key)
        if "bitRate" not in row or row["bitRate"] is None:
            raise V81Error("Application summary API response is missing bitRate; "
                           "no calculated fallback will be used.")
        metadata = application_metadata(lookup, row.get("appId_String"))
        result.append({
            "Rank": rank,
            "Application Name": metadata.get("application_short_name") or row.get("appId_String"),
            "Application Type Code": metadata.get("application_type_code"),
            "Application Type": metadata.get("application_type"),
            "Parent Application": metadata.get("parent_protocol"),
            "Application Group": metadata.get("application_group"),
            "App ID": row.get("appId"),
            STEP2_BIT_RATE_COLUMN: row["bitRate"],
            "Bytes": row.get("octets"),
            "Packets": row.get("packets"),
            "Device Name": interface.get("deviceName"),
            "Device IP": interface.get("deviceIPAddress"),
            "Interface Name": interface.get("interfaceName"),
            "Interface Alias": interface.get("alias"),
            "Interface Number": interface.get("interfaceNumber"),
        })
    return result


def run_step2(
    settings: NetscoutSettings,
    device_name: str,
    interface_filter: str,
    duration: str,
    top_n: int,
    output_dir: str | Path = ".",
    output_file: str | None = None,
) -> StepResult:
    """STEP 2: Device / Interface 기준 Application Bit Rate Top N을 생성한다."""
    time_def = build_time_def(duration, None, None, "NO_RESOLUTION")
    return run_step2_for_time_def(settings, device_name, interface_filter, time_def, top_n, output_dir, output_file)


def run_step2_for_time_def(
    settings: NetscoutSettings,
    device_name: str,
    interface_filter: str,
    time_def: TimeDef,
    top_n: int,
    output_dir: str | Path = ".",
    output_file: str | None = None,
) -> StepResult:
    """STEP 2: TimeDef 기반으로 Application Bit Rate Top N을 생성한다."""
    inventory = collect_inventory(settings)
    interface = select_device_interface(inventory, device_name, interface_filter)
    monitor_ip = str(interface.get("deviceIPAddress"))
    ifn = str(interface.get("interfaceNumber"))
    xml_text = build_application_summary_query_xml(monitor_ip, ifn, time_def, top_n)
    raw_rows = fetch_application_traffic_rows_with_xml(
        settings, monitor_ip, ifn, time_def_label(time_def), True, xml_text
    )
    rows = build_step2_api_rows(raw_rows, interface, inventory.applications)
    output_path = stage_output_dir(output_dir, "01_top") / (
        output_file
        or f"{safe_filename_part(device_name)}_{safe_interface_name(interface)}_Top{int(top_n)}_{safe_filename_part(time_def_label(time_def))}_{timestamp_text()}.xlsx"
    )
    path = write_xlsx(output_path, [("Application_BPS_TopN", STEP2_COLUMNS, rows)])
    append_run_summary(output_dir, f"top output={path} device={device_name} interface={interface_filter} rows={len(rows)}")
    return StepResult(len(rows), path)


def run_step3(
    settings: NetscoutSettings,
    device_name: str,
    interface_filter: str,
    application_name: str,
    duration: str,
    output_dir: str | Path = ".",
    output_file: str | None = None,
    limit: int | None = 1000,
    dedup_ip: bool = False,
) -> StepResult:
    """STEP 3: 선택 Application의 Source/Destination IP/Port 상세를 생성한다."""
    time_def = build_time_def(duration, None, None, "NO_RESOLUTION")
    return run_step3_for_time_def(settings, device_name, interface_filter, application_name, time_def, output_dir, output_file, limit, dedup_ip)


def run_step3_for_time_def(
    settings: NetscoutSettings,
    device_name: str,
    interface_filter: str,
    application_name: str,
    time_def: TimeDef,
    output_dir: str | Path = ".",
    output_file: str | None = None,
    limit: int | None = 1000,
    dedup_ip: bool = False,
) -> StepResult:
    """STEP 3: TimeDef 기반으로 선택 Application 상세를 생성한다."""
    inventory = collect_inventory(settings)
    interface = select_device_interface(inventory, device_name, interface_filter)
    app_filter = find_application_id(inventory.applications, application_name)
    raw_rows = normalize_traffic_rows(
        fetch_application_traffic_rows_for_time_def(
            settings,
            str(interface.get("deviceIPAddress")),
            str(interface.get("interfaceNumber")),
            time_def,
            use_topn=False,
            top_n=0,
            application=app_filter,
        ),
        str(interface.get("deviceIPAddress")),
        str(interface.get("interfaceNumber")),
    )
    filtered_rows = filter_application_rows(raw_rows, application_name, limit)
    original_filtered_rows = list(filtered_rows)
    before_dedup_rows = build_step3_rows(original_filtered_rows, interface)
    if dedup_ip:
        filtered_rows = dedup_ip_pairs(filtered_rows)
    rows = build_step3_rows(filtered_rows, interface)
    application_dir = stage_output_dir(output_dir, "03_detail") / safe_filename_part(application_name)
    output_path = application_dir / (
        output_file
        or f"{safe_filename_part(device_name)}_{safe_interface_name(interface)}_{safe_filename_part(application_name)}_Detail_{safe_filename_part(time_def_label(time_def))}_{timestamp_text()}.xlsx"
    )
    sheets = [("Application_Detail", STEP3_COLUMNS, before_dedup_rows)]
    if dedup_ip:
        sheets.extend(
            [
                ("Application_Detail_After_Dedup", STEP3_COLUMNS, rows),
                ("Unique_IP_Pairs", UNIQUE_IP_PAIR_COLUMNS, build_unique_ip_pair_rows(original_filtered_rows, interface)),
                ("Unique_IP_Port_Pairs", UNIQUE_IP_PORT_PAIR_COLUMNS, build_unique_ip_port_pair_rows(original_filtered_rows, interface)),
            ]
        )
    path = write_xlsx(output_path, sheets)
    append_run_summary(output_dir, f"detail output={path} application={application_name} rows={len(rows)}")
    return StepResult(len(rows), path)


def filter_market_data_feed_applications(applications: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Market Data Feeds 계열 Application만 별도 확인용으로 추출한다."""
    result = []
    for application in applications:
        group = str(application.get("application_group") or "").strip().lower()
        app_type = str(application.get("application_type") or "").strip().lower()
        if group == "market data feeds" or app_type == "market data feeds":
            result.append(application)
    return result


def format_application_inventory_rows(applications: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """정규화된 Application row를 문서 표기와 가까운 Inventory 표시 컬럼으로 변환한다."""
    rows = []
    for application in applications:
        rows.append(
            {
                "ApplicationShortName": application.get("application_short_name"),
                "ApplicationLongName": application.get("application_long_name"),
                "ApplicationType": application.get("application_type"),
                "ApplicationTypeCode": application.get("application_type_code"),
                "ParentApplication": application.get("parent_protocol"),
                "ApplicationPort": application.get("application_port"),
                "AdditionalPort": application.get("additional_port"),
                "ApplicationGroup": application.get("application_group"),
                "IsActive": application.get("is_active"),
                "Logged": application.get("logged"),
                "ResponseTime": application.get("response_time"),
                "Session": application.get("session"),
                "RecordingType": application.get("recording_type"),
            }
        )
    return rows


def summarize_applications(rows: list[dict[str, Any]], top_n: int = DEFAULT_TOP_N) -> list[dict[str, Any]]:
    """Application 기준으로 bitRate / Byte / Packet을 합산한다."""
    grouped: dict[tuple[object, object], dict[str, Any]] = {}
    for row in rows:
        key = (row.get("appId_String"), row.get("appId"))
        if key not in grouped:
            grouped[key] = {
                "rank": None,
                "Application": row.get("appId_String"),
                "App ID": row.get("appId"),
                "Total bitRate": 0,
                "Total Mbps": 0,
                "Total bitRate In": 0,
                "Total bitRate Out": 0,
                "Total Byte": 0,
                "Total Byte In": 0,
                "Total Byte Out": 0,
                "Total Packets": 0,
                "Total Packets In": 0,
                "Total Packets Out": 0,
                "Flow Count": 0,
            }
        item = grouped[key]
        item["Total bitRate"] = sum_number(item["Total bitRate"], row.get("bitRate"))
        item["Total bitRate In"] = sum_number(item["Total bitRate In"], row.get("bitRateIn"))
        item["Total bitRate Out"] = sum_number(item["Total bitRate Out"], row.get("bitRateOut"))
        item["Total Byte In"] = sum_number(item["Total Byte In"], row.get("srcToDestOctets"))
        item["Total Byte Out"] = sum_number(item["Total Byte Out"], row.get("destToSrcOctets"))
        item["Total Packets In"] = sum_number(item["Total Packets In"], row.get("srcToDestPackets"))
        item["Total Packets Out"] = sum_number(item["Total Packets Out"], row.get("destToSrcPackets"))
        item["Flow Count"] += 1
    result = []
    for item in grouped.values():
        item["Total Mbps"] = (to_float(item["Total bitRate"]) or 0) / 1_000_000
        item["Total Byte"] = sum_number(item["Total Byte In"], item["Total Byte Out"])
        item["Total Packets"] = sum_number(item["Total Packets In"], item["Total Packets Out"])
        result.append(item)
    result = sorted(result, key=lambda item: to_float(item.get("Total bitRate")) or 0, reverse=True)[:top_n]
    for rank, item in enumerate(result, start=1):
        item["rank"] = rank
    return result


def filter_undefined_candidates(summary_rows: list[dict[str, Any]], min_bitrate: float, top_n: int) -> list[dict[str, Any]]:
    """OTHER/UNKNOWN 계열 중 min_bitrate 이상인 최종 확인 후보를 고른다."""
    candidates = []
    for row in summary_rows:
        app = str(row.get("Application") or "").upper()
        bitrate = to_float(row.get("Total bitRate")) or 0
        if bitrate < min_bitrate or not any(keyword in app for keyword in UNDEFINED_KEYWORDS):
            continue
        candidates.append(
            {
                "rank": None,
                "Application": row.get("Application"),
                "App ID": row.get("App ID"),
                "Total bitRate": row.get("Total bitRate"),
                "Total Mbps": row.get("Total Mbps"),
                "Total Byte": row.get("Total Byte"),
                "Flow Count": row.get("Flow Count"),
                "candidate_reason": "OTHER/UNKNOWN 계열 미식별 Application",
            }
        )
    candidates = candidates[:top_n]
    for rank, row in enumerate(candidates, start=1):
        row["rank"] = rank
    return candidates


def registered_application_keys(applications: list[dict[str, Any]]) -> set[str]:
    """등록 Application 판정에 사용할 이름/code 집합을 만든다."""
    keys: set[str] = set()
    for app in applications:
        for field in ("application_type_code", "application_short_name", "application_long_name"):
            value = app.get(field)
            if value not in (None, ""):
                keys.add(str(value).strip().upper())
    return keys


def application_lookup(applications: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """TopN 결과의 Application 이름으로 Config API Application 메타데이터를 찾기 위한 lookup을 만든다."""
    lookup: dict[str, dict[str, Any]] = {}
    for app in applications:
        for field in ("application_type_code", "application_short_name", "application_long_name"):
            value = app.get(field)
            if value not in (None, ""):
                lookup.setdefault(str(value).strip().upper(), app)
    return lookup


def application_metadata(lookup: dict[str, dict[str, Any]], application_name: Any) -> dict[str, Any]:
    """appId_String/Application 이름에 대응하는 Application List row를 반환한다."""
    return lookup.get(str(application_name or "").strip().upper(), {})


def classify_top_applications(
    app_summary: list[dict[str, Any]],
    applications: list[dict[str, Any]],
    drilldown_apps: list[str] | tuple[str, ...] = DEFAULT_DRILLDOWN_APPS,
) -> list[dict[str, Any]]:
    """TopN Application을 DRILLDOWN 우선순위로 분류한다."""
    registered = registered_application_keys(applications)
    lookup = application_lookup(applications)
    drilldown = {str(app).strip().upper() for app in drilldown_apps}
    rows = []
    for rank, row in enumerate(app_summary, start=1):
        name = str(row.get("Application") or "").strip()
        key = name.upper()
        metadata = application_metadata(lookup, name)
        if key in drilldown:
            status = "DRILLDOWN"
            action = drilldown_action_for_app(key)
            reason = f"{name} requires drilldown"
        elif key in registered:
            status = "EXISTING"
            action = "EXCLUDE"
            reason = "Existing application"
        elif name:
            status = "NEW"
            action = "REGISTER_CANDIDATE"
            reason = "TopN application not found in registered inventory"
        else:
            status = "REVIEW"
            action = "MANUAL_REVIEW"
            reason = "Application name is missing"
        rows.append(
            {
                "Rank": rank,
                "Application": name,
                "Application Type": metadata.get("application_type") or "",
                "Application Type Code": metadata.get("application_type_code") or "",
                "Parent Application": metadata.get("parent_protocol") or "",
                "Application Group": metadata.get("application_group") or "",
                "Traffic": row.get("Total bitRate") or row.get("Total Byte"),
                "Packets": row.get("Total Packets"),
                "Source": "TOPN",
                "Defined": "",
                "Status": status,
                "Action": action,
                "Reason": reason,
                "App ID": row.get("App ID"),
                "Application Key": "",
                "Port": "",
            }
        )
    return rows


def drilldown_action_for_app(app_name: str) -> str:
    if app_name in {"HTTPS", "HTTP", "QUIC"}:
        return "URL_DISCOVERY"
    if app_name in {"TCP_OTHER", "UDP_OTHER"}:
        return "PORT_DISCOVERY"
    return "MANUAL_REVIEW"


def drilldown_web(settings: NetscoutSettings) -> list[dict[str, Any]]:
    """URL/SNI/Web Application Discovery 결과를 신규 후보 형식으로 변환한다."""
    rows: list[dict[str, Any]] = []
    for endpoint, source in (
        ("/ng1api/ncm/asitrafficdiscovery/url", "URL_DISCOVERY"),
        ("/ng1api/ncm/asitrafficdiscovery/webapplications", "WEB_APPLICATION_DISCOVERY"),
    ):
        try:
            discovered = config_get_collection(settings, endpoint, f"Config API {source}")
        except Exception as exc:  # noqa: BLE001
            logger.warning("Skipping failed %s: %s", source, exc)
            continue
        for item in discovered:
            if source == "URL_DISCOVERY":
                rows.append(normalize_url_discovery_candidate(item))
            else:
                rows.append(normalize_web_application_candidate(item))
    return rows


def drilldown_ports(settings: NetscoutSettings, duration: str) -> list[dict[str, Any]]:
    """Port Discovery 결과를 신규 후보 형식으로 변환한다."""
    query = urlencode({"duration": duration}) if duration else ""
    endpoint = "/ng1api/ncm/asitrafficdiscovery/port" + (f"?{query}" if query else "")
    try:
        discovered = config_get_collection(settings, endpoint, "Config API Port Discovery")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Skipping failed Port Discovery: %s", exc)
        return []
    return [normalize_port_discovery_candidate(item) for item in discovered]


def normalize_url_discovery_candidate(item: dict[str, Any]) -> dict[str, Any]:
    name = first_value(item, "URLAddress", "urlAddress", "Definition", "definition")
    defined = to_bool(first_value(item, "defined", "Defined"))
    url_type = first_value(item, "URLType", "urlType")
    parent = parent_from_url_type(url_type)
    return candidate_row(
        application=name,
        parent=parent,
        traffic=first_value(item, "noOfHits", "NumOfHits", "TotalOctets"),
        packets=first_value(item, "TotalPackets", "packets"),
        source="URL_DISCOVERY",
        defined=defined,
        status="EXISTING" if defined else "NEW",
        action="EXCLUDE" if defined else "REGISTER_CANDIDATE",
        app_type=url_type,
        app_group=first_value(item, "ApplicationGroup", "applicationGroup"),
        reason="Existing discovered URL" if defined else f"Undefined {url_type or 'URL'}",
        app_key=first_value(item, "ApplicationKey", "applicationKey"),
    )


def normalize_web_application_candidate(item: dict[str, Any]) -> dict[str, Any]:
    name = first_value(item, "ApplicationShortName", "applicationShortName", "ApplicationKey", "applicationKey")
    defined = bool(first_value(item, "ApplicationTypeCode", "applicationTypeCode", "Definitions", "definitions"))
    return candidate_row(
        application=name,
        parent="WEB",
        traffic=first_value(item, "TotalOctets", "totalOctets", "NumOfHits"),
        packets=first_value(item, "TotalPackets", "totalPackets"),
        source="WEB_APPLICATION_DISCOVERY",
        defined=defined,
        status="EXISTING" if defined else "NEW",
        action="EXCLUDE" if defined else "REGISTER_CANDIDATE",
        app_type=first_value(item, "ApplicationTypeCode", "applicationTypeCode"),
        app_group="",
        reason="Existing web application" if defined else "Undefined web application",
        app_key=first_value(item, "ApplicationKey", "applicationKey"),
    )


def normalize_port_discovery_candidate(item: dict[str, Any]) -> dict[str, Any]:
    port = first_value(item, "port", "Port")
    parent = first_value(item, "ParentApplication", "parentApplication", "protocol", "Protocol")
    name = f"{parent or 'PORT'}-{port}" if port not in (None, "") else str(first_value(item, "ApplicationShortName", "applicationShortName") or "")
    defined = to_bool(first_value(item, "defined", "Defined"))
    return candidate_row(
        application=name,
        parent=parent,
        traffic=first_value(item, "volume", "octets", "TotalOctets"),
        packets=first_value(item, "packets", "Packets"),
        source="PORT_DISCOVERY",
        defined=defined,
        status="EXISTING" if defined else "NEW",
        action="EXCLUDE" if defined else "REGISTER_CANDIDATE",
        app_type=first_value(item, "ApplicationShortName", "applicationShortName"),
        app_group=first_value(item, "ApplicationGroup", "applicationGroup"),
        reason="Existing discovered port" if defined else "Undefined discovered port",
        port=port,
    )


def candidate_row(
    application: Any,
    parent: Any,
    traffic: Any,
    packets: Any,
    source: str,
    defined: Any,
    status: str,
    action: str,
    app_type: Any,
    app_group: Any,
    reason: str,
    app_id: Any = "",
    app_key: Any = "",
    port: Any = "",
) -> dict[str, Any]:
    return {
        "Rank": None,
        "Application": application,
        "Application Type": app_type,
        "Application Type Code": "",
        "Parent Application": parent,
        "Application Group": app_group,
        "Traffic": traffic,
        "Packets": packets,
        "Source": source,
        "Defined": defined,
        "Status": status,
        "Action": action,
        "Reason": reason,
        "App ID": app_id,
        "Application Key": app_key,
        "Port": port,
    }


def parent_from_url_type(url_type: Any) -> str:
    text = str(url_type or "").upper()
    if "QUIC" in text:
        return "QUIC"
    if "SSL" in text or "HTTPS" in text:
        return "HTTPS"
    if "URL" in text or "HTTP" in text:
        return "HTTP"
    return ""


def deduplicate_candidates(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """source별 key 기준으로 후보 중복을 제거한다."""
    seen: set[tuple[Any, ...]] = set()
    result = []
    for row in rows:
        key = candidate_dedup_key(row)
        if key in seen:
            continue
        seen.add(key)
        result.append(row)
    return result


def candidate_dedup_key(row: dict[str, Any]) -> tuple[Any, ...]:
    source = row.get("Source")
    if source in {"URL_DISCOVERY", "WEB_APPLICATION_DISCOVERY"}:
        return (source, normalize_key(row.get("Application")), normalize_key(row.get("Application Key")))
    if source == "PORT_DISCOVERY":
        return (source, normalize_key(row.get("Parent Application")), normalize_key(row.get("Port")))
    return (source, normalize_key(row.get("Application")), normalize_key(row.get("App ID")))


def rank_candidates(rows: list[dict[str, Any]], target_candidates: int) -> list[dict[str, Any]]:
    ranked = sorted(rows, key=lambda row: to_float(row.get("Traffic")) or 0, reverse=True)
    if target_candidates > 0:
        ranked = ranked[:target_candidates]
    for rank, row in enumerate(ranked, start=1):
        row["Rank"] = rank
    return ranked


def split_candidate_sheets(rows: list[dict[str, Any]]) -> dict[str, list[dict[str, Any]]]:
    return {
        "Existing": [row for row in rows if row.get("Status") == "EXISTING"],
        "Drilldown": [row for row in rows if row.get("Status") == "DRILLDOWN"],
        "New_Candidates": [row for row in rows if row.get("Status") == "NEW"],
        "Review": [row for row in rows if row.get("Status") == "REVIEW"],
    }


def first_value(record: dict[str, Any], *keys: str) -> Any:
    lower = {str(key).lower(): value for key, value in record.items()}
    for key in keys:
        if key in record:
            return record[key]
        value = lower.get(key.lower())
        if value is not None:
            return value
    return None


def to_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"true", "1", "yes", "y"}


def normalize_key(value: Any) -> str:
    return str(value or "").strip().upper()


def analyze_raw_json(
    raw_json_path: str | Path,
    output_path: str | Path,
    monitor_ip: str = "",
    ifn: str = "",
    top_n: int = DEFAULT_TOP_N,
    min_bitrate: float = DEFAULT_MIN_BITRATE,
    application: str | None = None,
    limit: int | None = None,
) -> TrafficResult:
    """저장된 raw JSON을 분석해 V8 Excel을 만든다."""
    raw_rows = filter_application_rows(normalize_traffic_rows(load_raw_json(raw_json_path), monitor_ip, ifn), application, limit)
    app_summary = summarize_applications(raw_rows, top_n)
    candidates = filter_undefined_candidates(app_summary, min_bitrate, top_n)
    summary = [
        {"metric": "source_raw_json", "value": str(raw_json_path)},
        {"metric": "monitor_ip", "value": monitor_ip},
        {"metric": "ifn", "value": ifn},
        {"metric": "raw_rows", "value": len(raw_rows)},
        {"metric": "top_n", "value": top_n},
        {"metric": "min_bitrate", "value": min_bitrate},
        {"metric": "application_filter", "value": application or ""},
        {"metric": "limit", "value": limit or ""},
        {"metric": "generated_at", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ]
    path = export_traffic_xlsx(output_path, raw_rows, app_summary, candidates, summary)
    return TrafficResult(len(raw_rows), len(app_summary), len(candidates), path)


def run_candidate_analysis(
    settings: NetscoutSettings,
    device_name: str,
    interface_filter: str,
    duration: str,
    topn_initial: int = 50,
    topn_max: int = 200,
    target_candidates: int = 50,
    output_dir: str | Path = ".",
    output_file: str | None = None,
    drilldown_apps: list[str] | tuple[str, ...] = DEFAULT_DRILLDOWN_APPS,
) -> StepResult:
    """Traffic TopN 기반으로 기존 등록 제거와 Drilldown을 수행해 신규 후보 리포트를 만든다."""
    inventory = collect_inventory(settings)
    interface = select_device_interface(inventory, device_name, interface_filter)
    raw_rows = normalize_traffic_rows(
        fetch_application_traffic_rows(
            settings,
            str(interface.get("deviceIPAddress")),
            str(interface.get("interfaceNumber")),
            duration,
            use_topn=False,
            top_n=topn_max,
        ),
        str(interface.get("deviceIPAddress")),
        str(interface.get("interfaceNumber")),
    )
    app_summary = summarize_applications(raw_rows, topn_max)
    topn_rows = classify_top_applications(app_summary[:topn_initial], inventory.applications, drilldown_apps)
    drilldown_names = {normalize_key(row.get("Application")) for row in topn_rows if row.get("Status") == "DRILLDOWN"}
    discovery_rows: list[dict[str, Any]] = []
    if drilldown_names & {"HTTPS", "HTTP", "QUIC"}:
        discovery_rows.extend(drilldown_web(settings))
    if drilldown_names & {"TCP_OTHER", "UDP_OTHER"}:
        discovery_rows.extend(drilldown_ports(settings, duration))

    all_rows = deduplicate_candidates(topn_rows + discovery_rows)
    sheets = split_candidate_sheets(all_rows)
    new_candidates = rank_candidates(sheets["New_Candidates"], target_candidates)
    sheets["New_Candidates"] = new_candidates
    summary = [
        {"metric": "device", "value": device_name},
        {"metric": "interface", "value": interface_filter},
        {"metric": "duration", "value": duration},
        {"metric": "topn_initial", "value": topn_initial},
        {"metric": "topn_max", "value": topn_max},
        {"metric": "target_candidates", "value": target_candidates},
        {"metric": "topn_raw", "value": len(topn_rows)},
        {"metric": "existing", "value": len(sheets["Existing"])},
        {"metric": "drilldown", "value": len(sheets["Drilldown"])},
        {"metric": "new_candidates", "value": len(sheets["New_Candidates"])},
        {"metric": "review", "value": len(sheets["Review"])},
        {"metric": "write_api_called", "value": "false"},
    ]
    output_path = Path(output_dir) / (
        output_file
        or f"{safe_filename_part(device_name)}_{safe_interface_name(interface)}_Candidates_{safe_filename_part(duration)}_{timestamp_text()}.xlsx"
    )
    path = export_candidate_report(output_path, summary, topn_rows, sheets)
    return StepResult(len(sheets["New_Candidates"]), path)


def export_inventory_xlsx(inventory: Inventory, output_path: str | Path) -> Path:
    """Application / Device / Interface 목록 Excel을 저장한다."""
    application_rows = format_application_inventory_rows(inventory.applications)
    summary = [
        {"metric": "applications", "value": len(inventory.applications)},
        {"metric": "devices", "value": len(inventory.devices)},
        {"metric": "interfaces", "value": len(inventory.interfaces)},
        {"metric": "generated_at", "value": datetime.now().strftime("%Y-%m-%d %H:%M:%S")},
    ]
    return write_xlsx(
        output_path,
        [
            ("Application_List", APPLICATION_INVENTORY_COLUMNS, application_rows),
            ("Device_List", DEVICE_COLUMNS, inventory.devices),
            ("Interface_List", INTERFACE_COLUMNS, inventory.interfaces),
            ("Run_Summary", RUN_SUMMARY_COLUMNS, summary),
        ],
    )


def export_candidate_report(
    output_path: str | Path,
    summary: list[dict[str, Any]],
    topn_rows: list[dict[str, Any]],
    sheets: dict[str, list[dict[str, Any]]],
) -> Path:
    """신규 Application 후보 분석 결과를 Excel로 저장한다."""
    return write_xlsx(
        output_path,
        [
            ("Summary", RUN_SUMMARY_COLUMNS, summary),
            ("TopN_Raw", CANDIDATE_COLUMNS, topn_rows),
            ("Existing", CANDIDATE_COLUMNS, sheets["Existing"]),
            ("Drilldown", CANDIDATE_COLUMNS, sheets["Drilldown"]),
            ("New_Candidates", CANDIDATE_COLUMNS, sheets["New_Candidates"]),
            ("Review", CANDIDATE_COLUMNS, sheets["Review"]),
        ],
    )


def export_traffic_xlsx(
    output_path: str | Path,
    raw_rows: list[dict[str, Any]],
    app_summary: list[dict[str, Any]],
    candidates: list[dict[str, Any]],
    summary: list[dict[str, Any]],
) -> Path:
    """V8 트래픽 분석 Excel을 저장한다."""
    return write_xlsx(
        output_path,
        [
            ("Traffic_Raw", TRAFFIC_RAW_COLUMNS, raw_rows),
            ("Application_Top10", APPLICATION_SUMMARY_COLUMNS, app_summary),
            ("Undefined_Candidates", UNDEFINED_CANDIDATE_COLUMNS, candidates),
            ("Run_Summary", RUN_SUMMARY_COLUMNS, summary),
        ],
    )


def write_xlsx(path: str | Path, sheets: list[tuple[str, list[str], list[dict[str, Any]]]]) -> Path:
    """여러 sheet를 가진 XLSX 파일을 생성한다."""
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    first = True
    for title, columns, rows in sheets:
        worksheet = workbook.active if first else workbook.create_sheet()
        first = False
        worksheet.title = title[:31]
        worksheet.append(columns)
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        for row in rows:
            worksheet.append([row.get(column) for column in columns])
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column_cells in worksheet.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 60)
    workbook.save(output_path)
    return output_path


def resolve_output_path(output_dir: str | Path | None, output_file: str | None, prefix: str) -> Path:
    """출력 파일 경로를 만든다."""
    directory = Path(output_dir or ".")
    filename = output_file or f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return directory / filename


def staged_output_path(output_dir: str | Path | None, stage: str, output_file: str | None, prefix: str) -> Path:
    """v0.86 날짜/단계 폴더 구조에 맞는 output 파일 경로를 만든다."""
    filename = output_file or f"{prefix}_{timestamp_text()}.xlsx"
    return stage_output_dir(output_dir, stage) / filename


def stage_output_dir(output_dir: str | Path | None, stage: str) -> Path:
    """`output/YYYYMMDD/<stage>` 형태의 단계별 output 폴더를 반환한다."""
    directory = Path(output_dir or ".")
    if directory.name == stage:
        return directory
    if directory.name.isdigit() and len(directory.name) == 8:
        return directory / stage
    return daily_output_dir(directory) / stage


def date_root_from_output_dir(output_dir: str | Path | None) -> Path:
    """`output/YYYYMMDD` 날짜 root를 반환한다."""
    return daily_output_dir(output_dir)


def append_run_summary(output_dir: str | Path | None, message: str) -> None:
    """v0.86 날짜 root 아래 run_summary.txt에 실행 이력을 남긴다."""
    summary_dir = date_root_from_output_dir(output_dir)
    summary_path = summary_dir / "run_summary.txt"
    line = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {message}\n"
    try:
        summary_dir.mkdir(parents=True, exist_ok=True)
        with summary_path.open("a", encoding="utf-8") as file:
            file.write(line)
    except OSError as exc:
        logger.warning("Failed to write run summary path=%s error=%s", summary_path, exc)


def safe_interface_name(interface: dict[str, Any]) -> str:
    """파일명에 사용할 Interface 이름을 만들고, 이름이 없으면 IF번호를 사용한다."""
    name = interface.get("interfaceName")
    if name not in (None, ""):
        return safe_filename_part(str(name))
    return f"IF{safe_filename_part(str(interface.get('interfaceNumber') or 'unknown'))}"


def safe_filename_part(value: object) -> str:
    """Windows 파일명에 사용할 수 없는 문자를 안전한 `_`로 치환한다."""
    text = str(value or "").strip()
    for char in '\\/:*?"<>|':
        text = text.replace(char, "_")
    return text.replace(" ", "_") or "unknown"


def timestamp_text() -> str:
    """파일명용 현재 시각 문자열을 반환한다."""
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def date_text() -> str:
    """실행일 단위 output 폴더명 문자열을 반환한다."""
    return datetime.now().strftime("%Y%m%d")


def daily_output_dir(output_dir: str | Path | None) -> Path:
    """Detail 결과를 날짜 폴더 아래에 모으기 위한 output root를 반환한다.

    사용자가 `./output`을 넘기면 `./output/YYYYMMDD`를 사용한다.
    이미 마지막 경로가 `YYYYMMDD`이면 중복으로 날짜를 붙이지 않는다.
    """
    directory = Path(output_dir or ".")
    if directory.name.isdigit() and len(directory.name) == 8:
        return directory
    return directory / date_text()








def to_float(value: object) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def sum_number(left: object, right: object) -> float:
    return (to_float(left) or 0) + (to_float(right) or 0)
