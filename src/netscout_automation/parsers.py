"""NETSCOUT Config API Application 응답 parser와 normalizer."""

from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from typing import Any


APPLICATION_COLUMNS = [
    "application_type_code",
    "application_short_name",
    "application_long_name",
    "application_port",
    "additional_port",
    "application_group",
    "parent_protocol",
    "application_type",
    "logged",
    "response_time",
    "session",
    "recording_type",
    "is_active",
]

APPLICATION_FIELD_ALIASES = {
    "application_type_code": ("applicationTypeCode", "ApplicationTypeCode"),
    "application_short_name": ("applicationShortName", "ApplicationShortName"),
    "application_long_name": ("applicationLongName", "ApplicationLongName"),
    "application_port": ("applicationPort", "ApplicationPort"),
    "additional_port": ("additionalPort", "AdditionalPort"),
    "application_group": ("applicationGroup", "ApplicationGroup"),
    "parent_protocol": ("parentProtocolTypeCode", "ParentApplication", "parentApplication"),
    "application_type": ("applicationType", "ApplicationType"),
    "logged": ("logged", "Logged"),
    "response_time": ("responseTime", "ResponseTime"),
    "session": ("session", "Session"),
    "recording_type": ("recordingType", "RecordingType"),
    "is_active": ("isActive", "IsActive"),
}

DEVICE_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "alias",
    "status",
    "nG1ServerName",
    "deviceType",
    "deviceProtocolSettings",
    "activeInterfaces",
    "inActiveInterfaces",
    "description",
    "communicationProtocol",
    "UUID",
]

DEVICE_FIELD_ALIASES = {
    "deviceName": ("deviceName", "DeviceName"),
    "deviceIPAddress": ("deviceIPAddress", "DeviceIPAddress"),
    "alias": ("alias", "Alias"),
    "status": ("status", "Status"),
    "nG1ServerName": ("nG1ServerName",),
    "deviceType": ("deviceType", "DeviceType"),
    "deviceProtocolSettings": ("deviceProtocolSettings", "DeviceProtocolSettings"),
    "activeInterfaces": ("activeInterfaces", "ActiveInterfaces"),
    "inActiveInterfaces": ("inActiveInterfaces", "InActiveInterfaces", "inactiveInterfaces"),
    "description": ("description", "Description"),
    "communicationProtocol": ("communicationProtocol", "CommunicationProtocol"),
    "UUID": ("UUID", "uuid"),
}

INTERFACE_COLUMNS = [
    "deviceName",
    "deviceIPAddress",
    "interfaceName",
    "alias",
    "interfaceNumber",
    "portSpeed",
    "interfaceLinkType",
    "interfaceSpeed",
    "status",
    "alarmTemplateName",
    "virtualization",
    "activeInterfaces",
    "inactiveInterfaces",
    "interface_key",
]

INTERFACE_FIELD_ALIASES = {
    "interfaceName": ("interfaceName", "InterfaceName"),
    "alias": ("alias", "Alias"),
    "interfaceNumber": ("interfaceNumber", "InterfaceNumber"),
    "portSpeed": ("portSpeed", "PortSpeed"),
    "interfaceLinkType": ("interfaceLinkType", "InterfaceLinkType"),
    "interfaceSpeed": ("interfaceSpeed", "InterfaceSpeed"),
    "status": ("status", "Status"),
    "alarmTemplateName": ("alarmTemplateName", "AlarmTemplateName"),
    "virtualization": ("virtualization", "Virtualization"),
    "activeInterfaces": ("activeInterfaces", "ActiveInterfaces"),
    "inactiveInterfaces": ("inactiveInterfaces", "InActiveInterfaces", "inactiveInterfaces"),
}


class ParserError(ValueError):
    """API 응답을 파싱할 수 없을 때 발생한다."""


def parse_applications_response(response: Any) -> list[dict[str, Any]]:
    """JSON, XML, 이미 로드된 object 형태의 Config API Application List를 파싱한다."""
    if isinstance(response, dict):
        return parse_applications_json(response)
    if isinstance(response, list):
        return parse_applications_json(response)
    if isinstance(response, str):
        stripped = response.strip()
        if not stripped:
            return []
        if stripped.startswith("{") or stripped.startswith("["):
            try:
                return parse_applications_json(json.loads(stripped))
            except json.JSONDecodeError as exc:
                raise ParserError("Invalid JSON applications response.") from exc
        if stripped.startswith("<"):
            return parse_applications_xml(stripped)
        preview = stripped[:120].replace("\r", " ").replace("\n", " ")
        raise ParserError(
            "Unsupported applications response format. Expected JSON or XML. "
            f"Response starts with: {preview!r}"
        )
    raise ParserError(f"Unsupported applications response type: {type(response).__name__}")


def parse_applications_json(payload: dict[str, Any] | list[Any]) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        records = payload
    else:
        records = payload.get("applicationConfigurations", [])

    if records is None:
        return []
    if not isinstance(records, list):
        raise ParserError("Expected applicationConfigurations to be a list.")

    return [normalize_application(record) for record in records if isinstance(record, dict)]


def parse_applications_xml(payload: str) -> list[dict[str, Any]]:
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ParserError("Invalid XML applications response.") from exc

    if _local_name(root.tag) == "ApplicationConfiguration":
        elements = [root]
    else:
        elements = [element for element in root.iter() if _local_name(element.tag) == "ApplicationConfiguration"]

    return [normalize_application(_xml_element_to_dict(element)) for element in elements]


def normalize_application(record: dict[str, Any]) -> dict[str, Any]:
    """NETSCOUT field name variant를 안정적인 내부 column 이름으로 정규화한다."""
    normalized: dict[str, Any] = {}
    for target, aliases in APPLICATION_FIELD_ALIASES.items():
        normalized[target] = _first_present(record, aliases)
    return normalized


def parse_devices_response(response: Any) -> list[dict[str, Any]]:
    """JSON, XML, 이미 로드된 object 형태의 Device List를 파싱한다."""
    records = _parse_collection_response(response, "deviceConfigurations", "DeviceConfiguration")
    return [normalize_device(record) for record in records if isinstance(record, dict)]


def parse_device_interfaces_response(
    response: Any,
    device_name: str | None = None,
    device_ip_address: str | None = None,
) -> list[dict[str, Any]]:
    """JSON/XML Interface List를 파싱하고 device 식별 정보를 보강한다."""
    records = _parse_collection_response(response, "interfaceConfigurations", "InterfaceConfiguration")
    return [
        normalize_interface(record, device_name=device_name, device_ip_address=device_ip_address)
        for record in records
        if isinstance(record, dict)
    ]


def normalize_device(record: dict[str, Any]) -> dict[str, Any]:
    """Device 응답 field variant를 export용 column으로 정규화한다."""
    return {target: _first_present(record, aliases) for target, aliases in DEVICE_FIELD_ALIASES.items()}


def normalize_interface(
    record: dict[str, Any],
    device_name: str | None = None,
    device_ip_address: str | None = None,
) -> dict[str, Any]:
    """Interface 응답 field variant를 정규화하고 interface_key를 만든다."""
    values = {target: _first_present(record, aliases) for target, aliases in INTERFACE_FIELD_ALIASES.items()}
    values["deviceName"] = _first_present(record, ("deviceName", "DeviceName")) or device_name
    values["deviceIPAddress"] = _first_present(record, ("deviceIPAddress", "DeviceIPAddress")) or device_ip_address
    values["interface_key"] = build_interface_key(values.get("deviceIPAddress"), values.get("interfaceNumber"))
    normalized = {column: values.get(column) for column in INTERFACE_COLUMNS}
    return normalized


def build_interface_key(device_ip_address: object, interface_number: object) -> str:
    """Data API FlowFilter에 쓰는 deviceIPAddress와 InterfaceNumber 조합으로 key를 만든다."""
    return f"{device_ip_address}|{interface_number}"


def _parse_collection_response(response: Any, json_key: str, xml_item_name: str) -> list[Any]:
    if isinstance(response, list):
        return response
    if isinstance(response, dict):
        for key in (json_key, json_key[0].upper() + json_key[1:], "response", "items"):
            value = response.get(key)
            if isinstance(value, list):
                return value
        return [response]
    if isinstance(response, str):
        stripped = response.strip()
        if not stripped:
            return []
        if stripped.startswith("{") or stripped.startswith("["):
            try:
                return _parse_collection_response(json.loads(stripped), json_key, xml_item_name)
            except json.JSONDecodeError as exc:
                raise ParserError(f"Invalid JSON {json_key} response.") from exc
        if stripped.startswith("<"):
            return _parse_xml_collection(stripped, xml_item_name)
    raise ParserError(f"Unsupported {json_key} response format.")


def _parse_xml_collection(payload: str, xml_item_name: str) -> list[dict[str, Any]]:
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise ParserError(f"Invalid XML {xml_item_name} response.") from exc
    if _local_name(root.tag) == xml_item_name:
        elements = [root]
    else:
        elements = [element for element in root.iter() if _local_name(element.tag) == xml_item_name]
    return [_xml_element_to_dict(element) for element in elements]


def _first_present(record: dict[str, Any], aliases: tuple[str, ...]) -> Any:
    lower_lookup = {key.lower(): value for key, value in record.items()}
    for alias in aliases:
        if alias in record:
            return record[alias]
        lower_alias = alias.lower()
        if lower_alias in lower_lookup:
            return lower_lookup[lower_alias]
    return None


def _xml_element_to_dict(element: ET.Element) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for child in element:
        key = _local_name(child.tag)
        if len(child):
            result[key] = _xml_element_to_dict(child)
        else:
            result[key] = (child.text or "").strip()
    return result


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
