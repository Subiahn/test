"""NETSCOUT API 자동화 도구의 CLI 진입점.

수집 목적별로 subcommand를 분리한다.
Config API catalog export, raw Data API probe, TCP_OTHER report,
raw inspection, unified workbook은 서로 다른 명령으로 실행한다.
"""

from __future__ import annotations

import argparse
import logging
from datetime import datetime
from pathlib import Path

from .client import NetscoutClient, NetscoutClientError
from .config import ConfigError, load_settings
from .data_probe import DataProbeError, build_probe_preview, run_data_probe
from .exporters import export_application_list_xlsx
from .parsers import ParserError, parse_applications_response
from .raw_inspection import RawInspectionError, format_raw_inspection, inspect_raw_json_fields
from .registration_candidates import (
    RegistrationCandidatesError,
    build_dry_run_preview,
    generate_registration_candidates_report,
)
from .tcp_other import TcpOtherError, generate_tcp_other_report
from .unified_report import UnifiedReportError, generate_unified_report, resolve_unified_output_path


def build_parser() -> argparse.ArgumentParser:
    """최상위 parser와 지원하는 모든 subcommand를 구성한다."""
    parser = argparse.ArgumentParser(prog="netscout_automation")
    subparsers = parser.add_subparsers(dest="command", required=True)

    config_applications = subparsers.add_parser(
        "config-applications",
        help="Fetch Config API Application List and export it to XLSX.",
    )
    config_applications.add_argument("--config", required=True, help="Path to YAML config file.")
    config_applications.add_argument("--env", default=".env.local", help="Path to local environment file.")
    config_applications.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    data_probe = subparsers.add_parser(
        "data-probe",
        help="Send a Data API XML probe and save the raw response.",
    )
    data_probe.add_argument("--config", required=True, help="Path to YAML config file.")
    data_probe.add_argument("--env", default=".env.local", help="Path to local environment file.")
    data_probe.add_argument("--xml-file", required=True, help="Path to probe XML request file.")
    data_probe.add_argument("--accept", choices=["json", "xml", "csv"], default="json", help="Response format to request.")
    data_probe.add_argument("--output-dir", default=None, help="Override output directory.")
    data_probe.add_argument("--dry-run", action="store_true", help="Print request details without calling Data API.")
    data_probe.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    tcp_other_report = subparsers.add_parser(
        "tcp-other-report",
        help="Search and fetch TCP_OTHER APPLICATION data, then export it to XLSX.",
    )
    tcp_other_report.add_argument("--config", required=True, help="Path to YAML config file.")
    tcp_other_report.add_argument("--env", default=".env.local", help="Path to local environment file.")
    tcp_other_report.add_argument("--application", default="TCP_OTHER", help="Application name to filter by appId_String.")
    tcp_other_report.add_argument("--duration", default="LAST_HOUR", help="Data API duration value.")
    tcp_other_report.add_argument("--output", required=True, help="Output XLSX path.")
    tcp_other_report.add_argument(
        "--mode",
        choices=["application", "topnconv"],
        default="topnconv",
        help="Data API query mode. topnconv includes Source/Destination Port details.",
    )
    tcp_other_report.add_argument(
        "--summary-by-host",
        action="store_true",
        help="Add TCP_OTHER_Host_Summary sheet aggregated by source/destination host and port.",
    )
    tcp_other_report.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    inspect_raw = subparsers.add_parser(
        "inspect-raw",
        help="Inspect Data API raw JSON response fields for Port/Byte/Trx probe work.",
    )
    inspect_raw.add_argument("--raw-file", required=True, help="Path to output/raw/*.json saved by data-probe.")
    inspect_raw.add_argument(
        "--fields",
        nargs="*",
        default=None,
        help="Optional field list to inspect. Defaults to Port/Byte/Trx probe candidates.",
    )
    inspect_raw.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    unified_report = subparsers.add_parser(
        "unified-report",
        help="Create a workbook with Application catalog, interface map, TCP_OTHER detail, and enriched detail.",
    )
    unified_report.add_argument("--config", required=True, help="Path to YAML config file.")
    unified_report.add_argument("--env", default=".env.local", help="Path to local environment file.")
    unified_report.add_argument("--application", default="TCP_OTHER", help="Application name to collect.")
    unified_report.add_argument("--duration", default="LAST_HOUR", help="Data API duration value.")
    unified_report.add_argument("--output", default=None, help="Output XLSX path. Defaults to output/netscout_unified_report_YYYYMMDD_HHMMSS.xlsx.")
    unified_report.add_argument("--timestamp", action="store_true", help="Append timestamp suffix to the specified --output filename.")
    unified_report.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    registration_candidates = subparsers.add_parser(
        "registration-candidates",
        help="Create an Application registration candidate workbook from TOPNCONV traffic.",
    )
    registration_candidates.add_argument("--config", required=True, help="Path to YAML config file.")
    registration_candidates.add_argument("--env", default=".env.local", help="Path to local environment file.")
    registration_candidates.add_argument("--duration", default="LAST_HOUR", help="Data API duration value.")
    registration_candidates.add_argument("--top-n", type=int, default=10, help="Top N rows per interface.")
    registration_candidates.add_argument("--output", required=True, help="Output XLSX path.")
    registration_candidates.add_argument("--dry-run", action="store_true", help="Print request plan without calling APIs.")
    registration_candidates.add_argument(
        "--active-only",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Collect only Active/ACT devices and interfaces by default.",
    )
    registration_candidates.add_argument("--max-devices", type=int, default=None, help="Limit number of devices to query.")
    registration_candidates.add_argument("--max-interfaces", type=int, default=None, help="Limit number of interfaces to query.")
    registration_candidates.add_argument("--debug", action="store_true", help="Enable debug logging without secrets.")

    return parser


def main(argv: list[str] | None = None) -> int:
    """파싱된 CLI 명령 하나를 실행하고 process exit code를 반환한다."""
    parser = build_parser()
    args = parser.parse_args(argv)
    _configure_logging(debug=args.debug)

    if args.command == "config-applications":
        return run_config_applications(args.config, args.env)
    if args.command == "data-probe":
        return run_data_probe_command(
            config_path=args.config,
            env_path=args.env,
            xml_file=args.xml_file,
            accept=args.accept,
            output_dir=args.output_dir,
            dry_run=args.dry_run,
        )
    if args.command == "tcp-other-report":
        return run_tcp_other_report_command(
            config_path=args.config,
            env_path=args.env,
            application=args.application,
            duration=args.duration,
            output=args.output,
            mode=args.mode,
            summary_by_host=args.summary_by_host,
        )
    if args.command == "inspect-raw":
        return run_inspect_raw_command(args.raw_file, args.fields)
    if args.command == "unified-report":
        return run_unified_report_command(args.config, args.env, args.application, args.duration, args.output, args.timestamp)
    if args.command == "registration-candidates":
        return run_registration_candidates_command(
            config_path=args.config,
            env_path=args.env,
            duration=args.duration,
            top_n=args.top_n,
            output=args.output,
            dry_run=args.dry_run,
            active_only=args.active_only,
            max_devices=args.max_devices,
            max_interfaces=args.max_interfaces,
        )

    parser.error(f"Unknown command: {args.command}")
    return 2


def run_config_applications(config_path: str, env_path: str) -> int:
    """Config API Application catalog를 수집해 Application_List workbook으로 저장한다."""
    try:
        settings = load_settings(config_path, env_path)
        raw_response = NetscoutClient(settings).get_applications_raw()
        applications = parse_applications_response(raw_response)
        output_path = _timestamped_applications_path(settings.output_directory)
        export_application_list_xlsx(applications, output_path)
    except (ConfigError, NetscoutClientError, ParserError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    logging.getLogger(__name__).info("Exported %s applications to %s", len(applications), output_path)
    return 0


def run_unified_report_command(
    config_path: str,
    env_path: str,
    application: str,
    duration: str,
    output: str | None,
    timestamp: bool = False,
) -> int:
    """Config API와 Data API 결과를 모아 4개 sheet unified workbook을 생성한다."""
    try:
        settings = load_settings(config_path, env_path)
        output_path = resolve_unified_output_path(output, settings.output_directory, timestamp)
        result = generate_unified_report(settings, application, duration, output_path)
    except (ConfigError, UnifiedReportError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    logging.getLogger(__name__).info(
        "Exported unified report applications=%s interfaces=%s details=%s application_summary=%s host_summary=%s output=%s",
        result.applications_count,
        result.interface_count,
        result.detail_count,
        result.application_summary_count,
        result.host_summary_count,
        result.output_path,
    )
    return 0


def run_tcp_other_report_command(
    config_path: str,
    env_path: str,
    application: str,
    duration: str,
    output: str,
    mode: str,
    summary_by_host: bool = False,
) -> int:
    """APPLICATION 또는 TOPNCONV mode로 TCP_OTHER 전용 workbook을 생성한다."""
    try:
        settings = load_settings(config_path, env_path)
        result = generate_tcp_other_report(settings, application, duration, output, mode, summary_by_host=summary_by_host)
    except (ConfigError, TcpOtherError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    logging.getLogger(__name__).info(
        "Exported TCP_OTHER report rows=%s targets=%s succeeded=%s failed=%s output=%s",
        result.row_count,
        result.searched_targets,
        result.detail_success_count,
        result.detail_failure_count,
        result.output_path,
    )
    return 0


def run_registration_candidates_command(
    config_path: str,
    env_path: str,
    duration: str,
    top_n: int,
    output: str,
    dry_run: bool,
    active_only: bool,
    max_devices: int | None,
    max_interfaces: int | None,
) -> int:
    """Application 등록 후보 분석 workbook을 생성하거나 dry-run 계획을 출력한다."""
    try:
        settings = load_settings(config_path, env_path)
        if dry_run:
            preview = build_dry_run_preview(settings, duration, top_n, output)
            logging.getLogger(__name__).info("registration-candidates dry-run: no API request will be sent")
            logging.getLogger(__name__).info("endpoints=%s", ", ".join(preview.endpoints))
            logging.getLogger(__name__).info("duration=%s", preview.duration)
            logging.getLogger(__name__).info("top_n=%s", preview.top_n)
            logging.getLogger(__name__).info("output=%s", preview.output_path)
            logging.getLogger(__name__).info("active_only=%s", active_only)
            logging.getLogger(__name__).info("max_devices=%s", max_devices)
            logging.getLogger(__name__).info("max_interfaces=%s", max_interfaces)
            logging.getLogger(__name__).info("credentials_present=%s", preview.credentials_present)
            return 0

        result = generate_registration_candidates_report(
            settings,
            duration,
            top_n,
            output,
            active_only=active_only,
            max_devices=max_devices,
            max_interfaces=max_interfaces,
        )
    except (ConfigError, RegistrationCandidatesError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    logging.getLogger(__name__).info(
        "Exported registration candidates devices=%s interfaces=%s topn_rows=%s tcp_other_rows=%s candidates=%s output=%s",
        result.devices_count,
        result.interfaces_count,
        result.topn_rows_count,
        result.tcp_other_rows_count,
        result.registration_candidate_count,
        result.output_path,
    )
    return 0


def run_inspect_raw_command(raw_file: str, fields: list[str] | None) -> int:
    """네트워크 호출 없이 저장된 raw probe JSON 파일을 검사한다."""
    try:
        result = inspect_raw_json_fields(raw_file, fields)
    except (RawInspectionError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    print(format_raw_inspection(result))
    return 0


def run_data_probe_command(
    config_path: str,
    env_path: str,
    xml_file: str,
    accept: str,
    output_dir: str | None,
    dry_run: bool,
) -> int:
    """dry-run이 아니면 사용자가 지정한 XML probe를 전송한다."""
    try:
        settings = load_settings(config_path, env_path)
        preview = build_probe_preview(settings, xml_file, accept)
        if dry_run:
            logging.getLogger(__name__).info("Data API probe dry-run: no request will be sent")
            logging.getLogger(__name__).info("request host=%s", preview.host)
            logging.getLogger(__name__).info("endpoint=%s", preview.endpoint)
            logging.getLogger(__name__).info("verify_ssl=%s", preview.verify_ssl)
            logging.getLogger(__name__).info("content_type=%s", preview.content_type)
            logging.getLogger(__name__).info("accept_type=%s", preview.accept_type)
            logging.getLogger(__name__).info("xml_file=%s", preview.xml_file)
            return 0

        result = run_data_probe(settings, xml_file, accept, output_dir)
    except (ConfigError, DataProbeError, OSError) as exc:
        logging.getLogger(__name__).error("%s", exc)
        return 1

    logging.getLogger(__name__).info(
        "Data API probe completed status=%s content_type=%s saved=%s",
        result.status_code,
        result.content_type,
        result.output_path,
    )
    return 0


def _timestamped_applications_path(output_directory: str) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return Path(output_directory) / f"netscout_applications_{timestamp}.xlsx"


def _configure_logging(debug: bool = False) -> None:
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s - %(message)s",
    )
