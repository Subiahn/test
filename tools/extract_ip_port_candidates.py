"""Detail Excel에서 Application별 IP/Port 후보를 추출한다.

이미 생성된 Application Detail Excel을 후처리해서 hostname 조회용 IP 목록과
IP/Port 후보 목록을 만든다. API 호출은 하지 않는다.
"""

from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ModuleNotFoundError:
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

try:
    import yaml
except ModuleNotFoundError:
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    import yaml


DETAIL_SHEET = "Application_Detail"
AFTER_DEDUP_SHEET = "Application_Detail_After_Dedup"
TEMP_PREFIX = "~$"

IP_PORT_COLUMNS = [
    "Application",
    "Direction",
    "IP",
    "Port",
    "Port Status",
    "Needs Review",
    "Review Reason",
    "Row Count",
    "Total bitRate",
    "Total Bytes",
    "Total Packets",
    "Devices",
    "Interfaces",
    "Source Files",
]
UNKNOWN_ROW_COLUMNS = [
    "Application",
    "Source IP",
    "Source Port",
    "Destination IP",
    "Destination Port",
    "Port Status",
    "Review Reason",
    "bitRate",
    "Bytes",
    "Packets",
    "Device Name",
    "Interface Name",
    "Interface Number",
    "Source File",
]
UNIQUE_IP_COLUMNS = ["IP", "Application Count", "Applications", "Candidate Count", "Needs Review Count"]
UNIQUE_IP_PORT_COLUMNS = ["IP", "Port", "Application Count", "Applications", "Candidate Count", "Total bitRate", "Total Bytes", "Total Packets"]
RUN_SUMMARY_COLUMNS = ["metric", "value"]


class ExtractIpPortError(RuntimeError):
    """IP/Port 후보 추출 중 복구할 수 없는 오류가 발생했을 때 사용한다."""


def normalize_key(value: object) -> str:
    return str(value or "").strip().upper()


def to_number(value: object) -> float:
    if value in (None, ""):
        return 0.0
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return 0.0


def normalize_port(value: object) -> int | None:
    if value in (None, ""):
        return None
    try:
        return int(float(str(value).replace(",", "")))
    except ValueError:
        return None


def load_yaml(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise ExtractIpPortError(f"YAML file not found: {path}")
    with path.open("r", encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    if not isinstance(data, dict):
        raise ExtractIpPortError(f"YAML root must be a mapping: {path}")
    return data


def enabled_application_keys_from_targets(path: Path | None, include_disabled: bool = False) -> set[str]:
    """Detail targets YAML에서 처리 대상 Application key를 추출한다."""
    if not path:
        return set()
    data = load_yaml(path)
    keys: set[str] = set()
    for target in data.get("targets", []):
        if not isinstance(target, dict):
            continue
        if target.get("enabled") is False and not include_disabled:
            continue
        for field in ("application_type_code", "application", "Application Type Code", "Application Name"):
            key = normalize_key(target.get(field))
            if key:
                keys.add(key)
    return keys


def find_detail_files(detail_root: Path) -> list[Path]:
    if not detail_root.exists():
        raise ExtractIpPortError(f"Detail root not found: {detail_root}")
    return sorted(
        path
        for path in detail_root.rglob("*.xlsx")
        if path.is_file() and not path.name.startswith(TEMP_PREFIX)
    )


def read_detail_rows(path: Path) -> list[dict[str, Any]]:
    """Detail Excel에서 Application_Detail 계열 sheet를 dict row로 읽는다."""
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet_name = DETAIL_SHEET if DETAIL_SHEET in workbook.sheetnames else None
    if sheet_name is None and AFTER_DEDUP_SHEET in workbook.sheetnames:
        sheet_name = AFTER_DEDUP_SHEET
    if sheet_name is None:
        return []

    worksheet = workbook[sheet_name]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(header or "").strip() for header in rows[0]]
    result: list[dict[str, Any]] = []
    for values in rows[1:]:
        row = {headers[index]: values[index] if index < len(values) else None for index in range(len(headers)) if headers[index]}
        if not any(value not in (None, "") for value in row.values()):
            continue
        row["Source File"] = str(path)
        row["Source Application Folder"] = path.parent.name
        result.append(row)
    return result


def row_application(row: dict[str, Any]) -> str:
    return str(row.get("Application Name") or row.get("Application") or row.get("Source Application Folder") or "").strip()


def row_application_keys(row: dict[str, Any]) -> set[str]:
    """Detail row와 파일 위치에서 Application 매칭 후보를 만든다.

    Detail Excel의 `Application Name`은 실제 Data API의 appId_String이고,
    YAML의 `application`/`application_type_code`도 보통 이 값을 사용한다.
    그래도 일부 환경에서는 표시명과 코드가 어긋날 수 있으므로 파일이 들어 있는
    `output/YYYYMMDD/<Application>/...xlsx` 폴더명도 함께 매칭 후보로 사용한다.
    """
    keys = {
        normalize_key(row.get("Application Name")),
        normalize_key(row.get("Application")),
        normalize_key(row.get("Application Type Code")),
        normalize_key(row.get("Source Application Folder")),
    }
    return {key for key in keys if key}


def row_matches_enabled(row: dict[str, Any], enabled_keys: set[str]) -> bool:
    if not enabled_keys:
        return True
    return bool(row_application_keys(row) & enabled_keys)


def row_port_status(src_port: int | None, dest_port: int | None) -> tuple[str, bool, str]:
    """row 단위 port 상태를 해석한다."""
    if src_port is None and dest_port is None:
        return "missing", True, "source_port_and_destination_port_missing"
    src_zero = src_port in (None, 0)
    dest_zero = dest_port in (None, 0)
    if src_zero and dest_zero:
        return "both_unknown", True, "source_port_and_destination_port_are_zero"
    if src_zero or dest_zero:
        return "partial_known", False, "one_side_port_is_zero"
    return "known", False, ""


def candidate_status(candidate_port: int | None, row_status: str) -> tuple[str, bool, str]:
    if row_status in {"both_unknown", "missing"}:
        return row_status, True, "port_unknown"
    if candidate_port in (None, 0):
        return "unknown", True, "candidate_port_is_zero"
    return "known", False, ""


def make_candidate(row: dict[str, Any], direction: str, ip: object, port: int | None, row_status: str) -> dict[str, Any] | None:
    ip_text = str(ip or "").strip()
    if not ip_text:
        return None
    # 후보는 실제 조회/등록 판단에 쓸 수 있는 IP+Port 조합만 남긴다.
    # NETSCOUT Detail에서 port 0은 포트를 알 수 없다는 의미로 보고 후보에서 제외한다.
    if port in (None, 0):
        return None
    port_value = port
    status, needs_review, reason = candidate_status(port_value, row_status)
    return {
        "Application": row_application(row),
        "Direction": direction,
        "IP": ip_text,
        "Port": port_value,
        "Port Status": status,
        "Needs Review": needs_review,
        "Review Reason": reason,
        "bitRate": to_number(row.get("bitRate")),
        "Bytes": to_number(row.get("Bytes")),
        "Packets": to_number(row.get("Packets")),
        "Device Name": row.get("Device Name"),
        "Interface Name": row.get("Interface Name"),
        "Interface Number": row.get("Interface Number"),
        "Source File": row.get("Source File"),
        "Source Application Folder": row.get("Source Application Folder"),
    }


def extract_candidates(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Detail row 목록에서 source/destination IP/Port 후보와 확인 필요 원본 row를 만든다."""
    candidates: list[dict[str, Any]] = []
    unknown_rows: list[dict[str, Any]] = []
    for row in rows:
        src_port = normalize_port(row.get("Source Port"))
        dest_port = normalize_port(row.get("Destination Port"))
        row_status, needs_review, review_reason = row_port_status(src_port, dest_port)

        for candidate in (
            make_candidate(row, "source", row.get("Source IP"), src_port, row_status),
            make_candidate(row, "destination", row.get("Destination IP"), dest_port, row_status),
        ):
            if candidate:
                candidates.append(candidate)

        if needs_review:
            unknown_rows.append(
                {
                    "Application": row_application(row),
                    "Source IP": row.get("Source IP"),
                    "Source Port": 0 if src_port is None else src_port,
                    "Destination IP": row.get("Destination IP"),
                    "Destination Port": 0 if dest_port is None else dest_port,
                    "Port Status": row_status,
                    "Review Reason": review_reason,
                    "bitRate": to_number(row.get("bitRate")),
                    "Bytes": to_number(row.get("Bytes")),
                    "Packets": to_number(row.get("Packets")),
                    "Device Name": row.get("Device Name"),
                    "Interface Name": row.get("Interface Name"),
                    "Interface Number": row.get("Interface Number"),
                    "Source File": row.get("Source File"),
                }
            )
    return candidates, unknown_rows


def aggregate_candidates(candidates: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[Any, ...], dict[str, Any]] = {}
    for row in candidates:
        key = (row["Application"], row["Direction"], row["IP"], row["Port"], row["Port Status"])
        if key not in grouped:
            grouped[key] = {
                "Application": row["Application"],
                "Direction": row["Direction"],
                "IP": row["IP"],
                "Port": row["Port"],
                "Port Status": row["Port Status"],
                "Needs Review": bool(row["Needs Review"]),
                "Review Reason": row["Review Reason"],
                "Row Count": 0,
                "Total bitRate": 0.0,
                "Total Bytes": 0.0,
                "Total Packets": 0.0,
                "Devices": set(),
                "Interfaces": set(),
                "Source Files": set(),
            }
        item = grouped[key]
        item["Row Count"] += 1
        item["Total bitRate"] += to_number(row.get("bitRate"))
        item["Total Bytes"] += to_number(row.get("Bytes"))
        item["Total Packets"] += to_number(row.get("Packets"))
        if row.get("Device Name"):
            item["Devices"].add(str(row.get("Device Name")))
        interface_label = "|".join(str(row.get(part) or "") for part in ("Interface Name", "Interface Number")).strip("|")
        if interface_label:
            item["Interfaces"].add(interface_label)
        if row.get("Source File"):
            item["Source Files"].add(str(row.get("Source File")))

    result = []
    for item in grouped.values():
        copied = dict(item)
        copied["Devices"] = ", ".join(sorted(item["Devices"]))
        copied["Interfaces"] = ", ".join(sorted(item["Interfaces"]))
        copied["Source Files"] = ", ".join(sorted(Path(path).name for path in item["Source Files"]))
        result.append(copied)
    result.sort(key=lambda row: (str(row["Application"]), -to_number(row["Total bitRate"]), str(row["IP"]), int(row["Port"])))
    return result


def build_unique_ip_rows(candidate_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, Any]] = {}
    for row in candidate_rows:
        ip = str(row.get("IP") or "")
        if not ip:
            continue
        if ip not in grouped:
            grouped[ip] = {
                "IP": ip,
                "Application Count": 0,
                "Applications": set(),
                "Candidate Count": 0,
                "Needs Review Count": 0,
            }
        item = grouped[ip]
        item["Applications"].add(str(row.get("Application") or ""))
        item["Candidate Count"] += int(row.get("Row Count") or 0)
        if row.get("Needs Review"):
            item["Needs Review Count"] += int(row.get("Row Count") or 0)

    result = []
    for item in grouped.values():
        copied = dict(item)
        copied["Application Count"] = len(item["Applications"])
        copied["Applications"] = ", ".join(sorted(app for app in item["Applications"] if app))
        result.append(copied)
    result.sort(key=lambda row: str(row["IP"]))
    return result


def build_unique_ip_port_rows(candidate_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """IP+Port 기준으로 최종 중복 제거 row를 만든다.

    IP/Port 후보는 이미 port 0을 제외한 상태라서, 여기서는 같은 IP와 Port가
    여러 Application 또는 source/destination 방향에서 반복된 경우 하나로 합친다.
    """
    grouped: dict[tuple[str, int], dict[str, Any]] = {}
    for row in candidate_rows:
        ip = str(row.get("IP") or "").strip()
        port = normalize_port(row.get("Port"))
        if not ip or port in (None, 0):
            continue
        key = (ip, port)
        if key not in grouped:
            grouped[key] = {
                "IP": ip,
                "Port": port,
                "Application Count": 0,
                "Applications": set(),
                "Candidate Count": 0,
                "Total bitRate": 0.0,
                "Total Bytes": 0.0,
                "Total Packets": 0.0,
            }
        item = grouped[key]
        application = str(row.get("Application") or "").strip()
        if application:
            item["Applications"].add(application)
        item["Candidate Count"] += int(row.get("Row Count") or 0)
        item["Total bitRate"] += to_number(row.get("Total bitRate"))
        item["Total Bytes"] += to_number(row.get("Total Bytes"))
        item["Total Packets"] += to_number(row.get("Total Packets"))

    result = []
    for item in grouped.values():
        copied = dict(item)
        copied["Application Count"] = len(item["Applications"])
        copied["Applications"] = ", ".join(sorted(item["Applications"]))
        result.append(copied)
    result.sort(key=lambda row: (-to_number(row.get("Total bitRate")), str(row["IP"]), int(row["Port"])))
    return result


def write_sheet(workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet = workbook.create_sheet(sheet_name)
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


def write_xlsx(
    output_path: Path,
    candidate_rows: list[dict[str, Any]],
    unknown_rows: list[dict[str, Any]],
    unique_ip_rows: list[dict[str, Any]],
    unique_ip_port_rows: list[dict[str, Any]],
    summary_rows: list[dict[str, Any]],
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "IP_Port_Candidates", IP_PORT_COLUMNS, candidate_rows)
    write_sheet(workbook, "Unknown_Port_Rows", UNKNOWN_ROW_COLUMNS, unknown_rows)
    write_sheet(workbook, "Unique_IPs", UNIQUE_IP_COLUMNS, unique_ip_rows)
    write_sheet(workbook, "Unique_IP_Port", UNIQUE_IP_PORT_COLUMNS, unique_ip_port_rows)
    write_sheet(workbook, "Run_Summary", RUN_SUMMARY_COLUMNS, summary_rows)
    workbook.save(output_path)
    return output_path


def parse_application_filter(value: str | None) -> set[str]:
    """쉼표로 구분한 Application 이름 필터를 정규화한다."""
    if not value:
        return set()
    return {str(item).strip().upper() for item in value.split(",") if str(item).strip()}


def filter_rows_by_applications(rows: list[dict[str, Any]], applications: set[str]) -> list[dict[str, Any]]:
    """Application 필터가 지정된 경우 해당 Application row만 남긴다."""
    if not applications:
        return rows
    return [row for row in rows if normalize_key(row.get("Application")) in applications or normalize_key(row.get("Source Application Folder")) in applications]


def write_yaml_output(
    candidate_rows: list[dict[str, Any]],
    unique_ip_rows: list[dict[str, Any]],
    unique_ip_port_rows: list[dict[str, Any]],
    output_path: Path,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    applications: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in candidate_rows:
        applications[str(row.get("Application") or "")].append(
            {
                "direction": row.get("Direction"),
                "ip": row.get("IP"),
                "port": row.get("Port"),
                "port_status": row.get("Port Status"),
                "needs_review": bool(row.get("Needs Review")),
                "row_count": row.get("Row Count"),
                "total_bitrate": row.get("Total bitRate"),
                "total_bytes": row.get("Total Bytes"),
                "total_packets": row.get("Total Packets"),
            }
        )
    payload = {
        "description": "Application Detail Excel에서 추출한 hostname 조회 및 등록 검토용 IP/Port 후보.",
        "unique_ip_ports": [
            {
                "ip": row["IP"],
                "port": row["Port"],
                "applications": [app.strip() for app in str(row.get("Applications") or "").split(",") if app.strip()],
                "candidate_count": row.get("Candidate Count"),
                "total_bitrate": row.get("Total bitRate"),
                "total_bytes": row.get("Total Bytes"),
                "total_packets": row.get("Total Packets"),
            }
            for row in unique_ip_port_rows
        ],
        "applications": [
            {"application": application, "candidates": rows}
            for application, rows in sorted(applications.items())
        ],
    }
    with output_path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(payload, file, allow_unicode=True, sort_keys=False)
    return output_path


def extract_ip_port_report(
    detail_root: Path,
    output: Path,
    targets_yaml: Path | None = None,
    yaml_output: Path | None = None,
    include_disabled: bool = False,
    applications: str | None = None,
) -> tuple[Path, Path | None, dict[str, int]]:
    enabled_keys = enabled_application_keys_from_targets(targets_yaml, include_disabled)
    application_filter = parse_application_filter(applications)
    all_rows: list[dict[str, Any]] = []
    detail_files = find_detail_files(detail_root)
    for path in detail_files:
        for row in read_detail_rows(path):
            if row_matches_enabled(row, enabled_keys):
                all_rows.append(row)

    candidates, unknown_rows = extract_candidates(all_rows)
    candidate_rows = aggregate_candidates(candidates)
    candidate_rows = filter_rows_by_applications(candidate_rows, application_filter)
    unknown_rows = filter_rows_by_applications(unknown_rows, application_filter)
    unique_ip_rows = build_unique_ip_rows(candidate_rows)
    unique_ip_port_rows = build_unique_ip_port_rows(candidate_rows)
    summary_rows = [
        {"metric": "detail_files_found", "value": len(detail_files)},
        {"metric": "detail_rows_read", "value": len(all_rows)},
        {"metric": "candidate_rows", "value": len(candidate_rows)},
        {"metric": "unique_ips", "value": len(unique_ip_rows)},
        {"metric": "unique_ip_ports", "value": len(unique_ip_port_rows)},
        {"metric": "unknown_port_rows", "value": len(unknown_rows)},
        {"metric": "targets_yaml_filter_enabled", "value": bool(enabled_keys)},
        {"metric": "application_filter", "value": ", ".join(sorted(application_filter))},
    ]
    xlsx_path = write_xlsx(output, candidate_rows, unknown_rows, unique_ip_rows, unique_ip_port_rows, summary_rows)
    yaml_path = write_yaml_output(candidate_rows, unique_ip_rows, unique_ip_port_rows, yaml_output) if yaml_output else None
    stats = {str(row["metric"]): int(row["value"]) if isinstance(row["value"], int) else row["value"] for row in summary_rows}
    return xlsx_path, yaml_path, stats


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract IP/Port candidates from NETSCOUT Detail Excel files.")
    parser.add_argument("--detail-root", required=True, type=Path, help="Application Detail Excel들이 저장된 output root 경로.")
    parser.add_argument("--targets-yaml", type=Path, default=None, help="enabled Application만 처리할 detail targets YAML 경로.")
    parser.add_argument("--output", required=True, type=Path, help="IP/Port 후보 Excel 저장 경로.")
    parser.add_argument("--yaml-output", type=Path, default=None, help="hostname 조회용 YAML 저장 경로.")
    parser.add_argument("--include-disabled", action="store_true", help="targets YAML에서 enabled=false인 Application도 포함한다.")
    parser.add_argument("--applications", default=None, help="YAML/후보 생성 대상 Application 목록. 예: HTTPS,TCP_OTHER")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        xlsx_path, yaml_path, stats = extract_ip_port_report(
            args.detail_root,
            args.output,
            args.targets_yaml,
            args.yaml_output,
            args.include_disabled,
            args.applications,
        )
    except ExtractIpPortError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    print(
        "completed "
        f"detail_files={stats['detail_files_found']} detail_rows={stats['detail_rows_read']} "
        f"candidate_rows={stats['candidate_rows']} unique_ips={stats['unique_ips']} "
        f"unknown_port_rows={stats['unknown_port_rows']} output={xlsx_path} yaml_output={yaml_path or ''}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
