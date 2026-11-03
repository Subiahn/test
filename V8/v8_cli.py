"""NETSCOUT Automation v0.89 CLI.

V8은 사람이 중간 결과를 확인하면서 다음 실행 대상을 고르는 구조다.
그래서 단일 실행 명령뿐 아니라 YAML 생성/읽기 명령도 CLI에 포함한다.
"""

from __future__ import annotations

import argparse
import importlib.util
import logging
import sys
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
# 배포 폴더 안에서 직접 실행해도 공통 src 패키지를 찾을 수 있게 한다.
for path in (PROJECT_ROOT, SRC_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from netscout_automation.config import ConfigError, load_settings

try:
    from .v8_core import (
        V81Error,
        append_run_summary,
        build_time_def,
        collect_inventory,
        export_inventory_xlsx,
        resolve_output_path,
        run_detail_from_targets_yaml,
        run_detail_targets_yaml,
        run_step1,
        run_step2_for_time_def,
        run_step3_for_time_def,
        run_targets_yaml,
        run_top_from_targets_yaml,
        stage_output_dir,
        staged_output_path,
    )
except ImportError:
    from v8_core import (
        V81Error,
        append_run_summary,
        build_time_def,
        collect_inventory,
        export_inventory_xlsx,
        resolve_output_path,
        run_detail_from_targets_yaml,
        run_detail_targets_yaml,
        run_step1,
        run_step2_for_time_def,
        run_step3_for_time_def,
        run_targets_yaml,
        run_top_from_targets_yaml,
        stage_output_dir,
        staged_output_path,
    )

logger = logging.getLogger(__name__)

DEFAULT_TARGETS_YAML = str(Path("config") / "targets.yaml")
DEFAULT_DETAIL_TARGETS_YAML = str(Path("config") / "detail_targets.yaml")


def configure_logging(debug: bool = False) -> None:
    logging.basicConfig(level=logging.DEBUG if debug else logging.INFO, format="%(asctime)s %(levelname)s %(name)s - %(message)s")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="NETSCOUT Automation v0.89. Inventory / TopN / Detail / Postprocess.")
    # 기존 작업 습관에 맞춰 `--top` 같은 플래그형 alias와 `top` 서브커맨드를 모두 지원한다.
    parser.add_argument("--debug", action="store_true", help="DEBUG 로그를 출력한다.")
    parser.add_argument("--inventory", action="store_true", help="Application / Device / Interface 전체 목록을 조회한다.")
    parser.add_argument("--top", action="store_true", help="Device / Interface 기준 Application Bit Rate Top N을 조회한다.")
    parser.add_argument("--detail", action="store_true", help="선택 Application의 Source/Destination IP / Port 상세를 조회한다.")
    parser.add_argument("--targets-yaml", action="store_true", help="Batch 실행 전에 사람이 편집할 Device / Interface 대상 YAML을 생성한다.")
    parser.add_argument("--run-targets", action="store_true", help="편집된 targets YAML을 읽어 enabled 대상의 Top N을 실행한다.")
    parser.add_argument("--detail-targets-yaml", action="store_true", help="Top N Excel을 읽어 Detail 실행 대상 YAML을 생성한다.")
    parser.add_argument("--run-detail-targets", action="store_true", help="편집된 Detail targets YAML을 읽어 enabled 대상의 Detail을 실행한다.")
    parser.add_argument("--postprocess", action="store_true", help="Run postprocess by application.")
    parser.add_argument("--extract-candidates", action="store_true", help="Extract IP/Port candidates from one Detail folder.")
    parser.add_argument("--hostname-lookup", action="store_true", help="Run hostname lookup from candidate YAML.")
    parser.add_argument("--compare-inventory-apps", action="store_true", help="Compare pure/current Inventory and create manual app whitelist.")
    parser.add_argument("--filter-detail-targets", action="store_true", help="Disable manually registered apps in Detail targets YAML.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    inventory = subparsers.add_parser("list", aliases=["inventory"], help="Application / Device / Interface 목록 Excel 생성")
    add_common(inventory)
    add_output(inventory)

    step1 = subparsers.add_parser("step1-list", help="STEP 1. NETSCOUT 현황 Application / Device / Interface 조회")
    add_common(step1)
    add_output(step1)

    step2 = subparsers.add_parser("step2-top", help="STEP 2. Device / Interface 기준 Application Bit Rate Top N 조회")
    add_common(step2)
    add_output(step2)
    step2.add_argument("--device", required=True, help="필수. Device Name.")
    step2.add_argument("--interface", required=True, help="필수. Interface Name 또는 Interface Number. 예: if3, 3")
    step2.add_argument("--duration", help="선택. NETSCOUT duration. 예: LAST_HOUR, LAST_24_HOURS, LAST_7_DAYS")
    step2.add_argument("--start-time", help="선택. 절대 시작 시간. 예: 2026-08-24_00:00:00")
    step2.add_argument("--end-time", help="선택. 절대 종료 시간. 예: 2026-08-31_00:00:00")
    step2.add_argument("--resolution", default="NO_RESOLUTION", help="선택. TimeDef resolution. 기본값 NO_RESOLUTION")
    step2.add_argument("--top-n", type=int, default=50, help="선택. Bit Rate 기준 Top N. 기본값 50.")

    top = subparsers.add_parser("top", help="Device / Interface 기준 Application Bit Rate Top N 조회")
    add_common(top)
    add_output(top)
    top.add_argument("--device", required=True, help="필수. Device Name.")
    top.add_argument("--interface", required=True, help="필수. Interface Name 또는 Interface Number. 예: if3, 3")
    top.add_argument("--duration", help="선택. NETSCOUT duration. 예: LAST_HOUR, LAST_24_HOURS, LAST_7_DAYS")
    top.add_argument("--start-time", help="선택. 절대 시작 시간. 예: 2026-08-24_00:00:00")
    top.add_argument("--end-time", help="선택. 절대 종료 시간. 예: 2026-08-31_00:00:00")
    top.add_argument("--resolution", default="NO_RESOLUTION", help="선택. TimeDef resolution. 기본값 NO_RESOLUTION")
    top.add_argument("--top-n", type=int, default=50, help="선택. Bit Rate 기준 Top N. 기본값 50.")

    step3 = subparsers.add_parser("step3-detail", help="STEP 3. 선택 Application 상세 IP / Port 조회")
    add_common(step3)
    add_output(step3)
    step3.add_argument("--device", required=True, help="필수. Device Name.")
    step3.add_argument("--interface", required=True, help="필수. Interface Name 또는 Interface Number. 예: if3, 3")
    step3.add_argument("--application", required=True, help="필수. Application Name. 예: BING, TCP_OTHER")
    step3.add_argument("--duration", help="선택. NETSCOUT duration. 예: LAST_HOUR, LAST_24_HOURS, LAST_7_DAYS")
    step3.add_argument("--start-time", help="선택. 절대 시작 시간. 예: 2026-08-24_00:00:00")
    step3.add_argument("--end-time", help="선택. 절대 종료 시간. 예: 2026-08-31_00:00:00")
    step3.add_argument("--resolution", default="NO_RESOLUTION", help="선택. TimeDef resolution. 기본값 NO_RESOLUTION")
    step3.add_argument("--limit", type=int, default=1000, help="선택. bitRate 내림차순 기준 상세 row 제한. 기본값 1000. 0 이하면 제한하지 않음.")
    step3.add_argument("--dedup-ip", action="store_true", help="선택. limit 적용 후 Source IP / Destination IP 조합 중복을 제거한다.")

    detail = subparsers.add_parser("detail", help="선택 Application 상세 IP / Port 조회")
    add_common(detail)
    add_output(detail)
    detail.add_argument("--device", required=True, help="필수. Device Name.")
    detail.add_argument("--interface", required=True, help="필수. Interface Name 또는 Interface Number. 예: if3, 3")
    detail.add_argument("--application", required=True, help="필수. Application Name. 예: BING, TCP_OTHER")
    detail.add_argument("--duration", help="선택. NETSCOUT duration. 예: LAST_HOUR, LAST_24_HOURS, LAST_7_DAYS")
    detail.add_argument("--start-time", help="선택. 절대 시작 시간. 예: 2026-08-24_00:00:00")
    detail.add_argument("--end-time", help="선택. 절대 종료 시간. 예: 2026-08-31_00:00:00")
    detail.add_argument("--resolution", default="NO_RESOLUTION", help="선택. TimeDef resolution. 기본값 NO_RESOLUTION")
    detail.add_argument("--limit", type=int, default=1000, help="선택. bitRate 내림차순 기준 상세 row 제한. 기본값 1000. 0 이하면 제한하지 않음.")
    detail.add_argument("--dedup-ip", action="store_true", help="선택. limit 적용 후 Source IP / Destination IP 조합 중복을 제거한다.")

    targets = subparsers.add_parser("targets-yaml", help="Batch 대상 Device / Interface YAML 생성")
    add_common(targets)
    targets.add_argument("--output", default=DEFAULT_TARGETS_YAML, help=f"선택. 생성할 YAML 파일 경로. 기본값 {DEFAULT_TARGETS_YAML}")
    targets.add_argument("--duration", default="LAST_7_DAYS", help="선택. 기본 duration. 기본값 LAST_7_DAYS")
    targets.add_argument("--top-n", type=int, default=50, help="선택. 기본 Top N. 기본값 50")
    targets.add_argument("--run-detail", action="store_true", help="선택. YAML 기본값으로 detail 자동 실행을 켠다.")
    targets.add_argument("--detail-top-n", type=int, default=10, help="선택. detail 대상 Top N 기본값. 기본값 10")
    targets.add_argument("--device-filter", help="선택. Device Name 부분 문자열 필터.")
    targets.add_argument("--interface-filter", help="선택. Interface Name 또는 Interface Number 필터. 예: if3, 3")

    run_targets = subparsers.add_parser("run-targets", help="targets YAML을 읽어 enabled 대상별 Top N 실행")
    add_common(run_targets)
    run_targets.add_argument("--targets", required=True, help="필수. targets-yaml 명령으로 만든 YAML 파일 경로.")
    run_targets.add_argument("--output-dir", default=".", help="선택. TopN Excel 저장 디렉터리.")
    run_targets.add_argument("--dry-run", action="store_true", help="선택. 실제 API 호출 없이 실행 예정 대상만 출력한다.")

    detail_targets = subparsers.add_parser("detail-targets-yaml", help="Top N Excel을 읽어 Detail 대상 YAML 생성")
    detail_targets.add_argument("--top-xlsx", required=True, help="필수. STEP 2 Top N Excel 파일 경로.")
    detail_targets.add_argument("--output", default=DEFAULT_DETAIL_TARGETS_YAML, help=f"선택. 생성할 Detail targets YAML 파일 경로. 기본값 {DEFAULT_DETAIL_TARGETS_YAML}")
    detail_targets.add_argument("--detail-top-n", type=int, default=50, help="선택. Rank 기준 Detail 대상 개수. 기본값 50.")
    detail_targets.add_argument("--limit", type=int, default=1000, help="선택. 각 Application Detail row 제한. 기본값 1000.")
    detail_targets.add_argument("--dedup-ip", action="store_true", help="선택. Detail 실행 시 Source/Destination IP 조합 중복 제거를 기본값으로 둔다.")

    run_detail_targets = subparsers.add_parser("run-detail-targets", help="Detail targets YAML을 읽어 enabled 대상별 Detail 실행")
    add_common(run_detail_targets)
    run_detail_targets.add_argument("--targets", required=True, help="필수. detail-targets-yaml 명령으로 만든 YAML 파일 경로.")
    run_detail_targets.add_argument("--output-dir", default=".", help="선택. Detail Excel 저장 디렉터리.")
    run_detail_targets.add_argument("--default-duration", default="LAST_7_DAYS", help="선택. YAML target에 duration이 비어 있을 때 사용할 기본 duration.")
    run_detail_targets.add_argument("--dry-run", action="store_true", help="선택. 실제 API 호출 없이 실행 예정 대상만 출력한다.")

    postprocess = subparsers.add_parser("postprocess", help="Postprocess Detail folders into IP/Port candidates")
    postprocess.add_argument("--date-root", required=True, type=Path, help="Required. output/YYYYMMDD date root.")
    postprocess.add_argument("--targets-yaml", type=Path, default=None, help="Optional. detail_targets.yaml or detail_targets_filtered.yaml.")
    postprocess.add_argument("--include-disabled", action="store_true", help="Include disabled targets.")
    postprocess.add_argument("--applications", default=None, help="Optional application filter. Example: HTTPS,TCP_OTHER")
    postprocess.add_argument("--dry-run", action="store_true", help="Show planned work without writing files.")
    postprocess.add_argument("--hostname-lookup", action="store_true", help="Run hostname lookup after candidate extraction.")
    postprocess.add_argument("--hostname-config", default="config/hostname_lookup.yaml", type=Path, help="Hostname lookup config YAML.")

    extract = subparsers.add_parser("extract-candidates", help="Extract IP/Port candidates from one Detail folder")
    extract.add_argument("--detail-root", required=True, type=Path, help="Required. Application Detail folder.")
    extract.add_argument("--targets-yaml", type=Path, default=None, help="Optional. detail targets YAML.")
    extract.add_argument("--output", required=True, type=Path, help="Required. candidate XLSX output path.")
    extract.add_argument("--yaml-output", type=Path, default=None, help="Optional. candidate YAML output path.")
    extract.add_argument("--include-disabled", action="store_true", help="Include disabled targets.")
    extract.add_argument("--applications", default=None, help="Optional application filter. Example: HTTPS,TCP_OTHER")

    hostname = subparsers.add_parser("hostname-lookup", help="Run hostname lookup from candidate YAML")
    hostname.add_argument("--input", required=True, type=Path, help="Required. ip_port_candidates.yaml path.")
    hostname.add_argument("--config", default="config/hostname_lookup.yaml", type=Path, help="Hostname lookup config YAML.")
    hostname.add_argument("--output", type=Path, help="Optional output XLSX path.")

    compare = subparsers.add_parser("compare-inventory-apps", help="Compare pure/current Inventory and create manual app whitelist")
    compare.add_argument("--pure-inventory", required=True, type=Path, help="Required. pure Inventory Excel path.")
    compare.add_argument("--current-inventory", required=True, type=Path, help="Required. current Inventory Excel path.")
    compare.add_argument("--output", required=True, type=Path, help="Required. CSV output path.")
    compare.add_argument("--yaml-output", type=Path, default=None, help="Optional. manual_registered_apps.yaml output path.")
    compare.add_argument("--sheet", default="Application_List", help="Inventory application sheet name.")
    compare.add_argument("--key-column", default="ApplicationTypeCode", help="Compare key column.")

    filter_targets = subparsers.add_parser("filter-detail-targets", help="Disable manually registered apps in Detail targets YAML")
    filter_targets.add_argument("--top-yaml", required=True, type=Path, help="Required. detail_targets.yaml path.")
    filter_targets.add_argument("--manual-apps-yaml", required=True, type=Path, help="Required. manual_registered_apps.yaml path.")
    filter_targets.add_argument("--output", required=True, type=Path, help="Required. filtered detail targets YAML output path.")
    filter_targets.add_argument("--disable-protected", action="store_true", help="Disable OTHER/UNKNOWN protection.")
    filter_targets.add_argument("--dry-run", action="store_true", help="Print stats without writing output.")

    return parser


def add_common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--config", required=True, help="필수. YAML config 경로.")
    parser.add_argument("--env", default=".env.local", help="선택. .env.local 경로.")


def add_output(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--output-dir", default=".", help="선택. 출력 디렉터리.")
    parser.add_argument("--output-file", help="선택. 출력 파일명.")


def main(argv: list[str] | None = None) -> int:
    argv = normalize_option_command(sys.argv[1:] if argv is None else argv)
    parser = build_parser()
    args = parser.parse_args(argv)
    configure_logging(args.debug)
    try:
        if args.command in {"list", "inventory"}:
            settings = load_settings(args.config, resolve_env_path(args.env))
            inventory = collect_inventory(settings)
            output = staged_output_path(args.output_dir, "00_inventory", args.output_file, "Netscout_Inventory")
            path = export_inventory_xlsx(inventory, output)
            append_run_summary(args.output_dir, f"inventory output={path} rows={len(inventory.applications) + len(inventory.devices) + len(inventory.interfaces)}")
            logger.info("v0.89 inventory completed applications=%s devices=%s interfaces=%s output=%s", len(inventory.applications), len(inventory.devices), len(inventory.interfaces), path)
            return 0
        if args.command == "step1-list":
            settings = load_settings(args.config, resolve_env_path(args.env))
            result = run_step1(settings, args.output_dir, args.output_file)
            logger.info("STEP 1 completed rows=%s output=%s", result.rows_count, result.output_path)
            return 0
        if args.command in {"step2-top", "top"}:
            settings = load_settings(args.config, resolve_env_path(args.env))
            time_def = build_time_def(args.duration, args.start_time, args.end_time, args.resolution)
            result = run_step2_for_time_def(settings, args.device, args.interface, time_def, args.top_n, args.output_dir, args.output_file)
            logger.info("STEP 2 completed rows=%s output=%s", result.rows_count, result.output_path)
            return 0
        if args.command in {"step3-detail", "detail"}:
            settings = load_settings(args.config, resolve_env_path(args.env))
            time_def = build_time_def(args.duration, args.start_time, args.end_time, args.resolution)
            result = run_step3_for_time_def(
                settings,
                args.device,
                args.interface,
                args.application,
                time_def,
                args.output_dir,
                args.output_file,
                args.limit,
                args.dedup_ip,
            )
            logger.info("STEP 3 completed rows=%s output=%s", result.rows_count, result.output_path)
            return 0
        if args.command == "targets-yaml":
            settings = load_settings(args.config, resolve_env_path(args.env))
            output = args.output
            if output == DEFAULT_TARGETS_YAML:
                output = stage_output_dir("output", "02_targets") / "targets.yaml"
            result = run_targets_yaml(
                settings,
                output,
                args.duration,
                args.top_n,
                args.run_detail,
                args.detail_top_n,
                args.device_filter,
                args.interface_filter,
            )
            logger.info("v0.89 targets YAML completed targets=%s output=%s", result.targets_count, result.output_path)
            return 0
        if args.command == "run-targets":
            settings = load_settings(args.config, resolve_env_path(args.env))
            result = run_top_from_targets_yaml(settings, args.targets, args.output_dir, args.dry_run)
            for row in result.results:
                logger.info(
                    "YAML target status=%s device=%s interface=%s duration=%s top_n=%s output=%s error=%s",
                    row["status"],
                    row["device"],
                    row["interface"],
                    row["duration"],
                    row["top_n"],
                    row["output"],
                    row["error"],
                )
            logger.info(
                "v0.89 run-targets completed planned=%s success=%s failed=%s skipped=%s dry_run=%s",
                result.planned_count,
                result.success_count,
                result.failure_count,
                result.skipped_count,
                args.dry_run,
            )
            return 0 if result.failure_count == 0 else 1
        if args.command == "detail-targets-yaml":
            # STEP 2 Excel만 읽어서 YAML을 만들기 때문에 config/env가 필요 없다.
            output = args.output
            if output == DEFAULT_DETAIL_TARGETS_YAML:
                output = default_detail_targets_output(args.top_xlsx)
            result = run_detail_targets_yaml(args.top_xlsx, output, args.detail_top_n, args.limit, args.dedup_ip)
            logger.info("v0.89 detail targets YAML completed targets=%s output=%s", result.targets_count, result.output_path)
            return 0
        if args.command == "run-detail-targets":
            settings = load_settings(args.config, resolve_env_path(args.env))
            result = run_detail_from_targets_yaml(settings, args.targets, args.output_dir, args.default_duration, args.dry_run)
            for row in result.results:
                logger.info(
                    "Detail YAML target status=%s rank=%s application=%s bps=%s device=%s interface=%s duration=%s limit=%s output=%s error=%s",
                    row["status"],
                    row["rank"],
                    row["application"],
                    row["bps"],
                    row["device"],
                    row["interface"],
                    row["duration"],
                    row["limit"],
                    row["output"],
                    row["error"],
                )
            logger.info(
                "v0.89 run-detail-targets completed planned=%s success=%s failed=%s skipped=%s dry_run=%s",
                result.planned_count,
                result.success_count,
                result.failure_count,
                result.skipped_count,
                args.dry_run,
            )
            return 0 if result.failure_count == 0 else 1
        if args.command == "postprocess":
            result = run_postprocess_command(args)
            for row in result["results"]:
                logger.info(
                    "postprocess application=%s status=%s output=%s yaml_output=%s candidate_rows=%s unique_ip_ports=%s unknown_port_rows=%s",
                    row.get("application", ""),
                    row.get("status", ""),
                    row.get("output", ""),
                    row.get("yaml_output", ""),
                    row.get("candidate_rows", ""),
                    row.get("unique_ip_ports", ""),
                    row.get("unknown_port_rows", ""),
                )
            logger.info(
                "v0.89 postprocess completed date_root=%s planned=%s skipped=%s hostname_lookup=%s",
                result.get("date_root", ""),
                result.get("planned_count", 0),
                result.get("skipped_count", 0),
                args.hostname_lookup,
            )
            return 0
        if args.command == "extract-candidates":
            module = load_tool_module("extract_ip_port_candidates.py", "netscout_auto_extract_ip_port_candidates_cli")
            xlsx_path, yaml_path, stats = module.extract_ip_port_report(
                args.detail_root,
                args.output,
                targets_yaml=args.targets_yaml,
                yaml_output=args.yaml_output,
                include_disabled=args.include_disabled,
                applications=args.applications,
            )
            logger.info(
                "v0.89 extract-candidates completed output=%s yaml_output=%s candidate_rows=%s unique_ip_ports=%s unknown_port_rows=%s",
                xlsx_path,
                yaml_path or "",
                stats.get("candidate_rows", 0),
                stats.get("unique_ip_ports", 0),
                stats.get("unknown_port_rows", 0),
            )
            return 0
        if args.command == "hostname-lookup":
            output = run_hostname_lookup(args.input, args.config, args.output)
            logger.info("v0.89 hostname-lookup completed output=%s", output)
            return 0
        if args.command == "compare-inventory-apps":
            module = load_tool_module("compare_inventory_apps.py", "netscout_auto_compare_inventory_apps_cli")
            csv_path, yaml_path, pure_count, current_count, diff_count = module.compare_inventory_apps(
                args.pure_inventory,
                args.current_inventory,
                args.output,
                args.yaml_output,
                args.sheet,
                args.key_column,
            )
            logger.info(
                "v0.89 compare-inventory-apps completed pure_rows=%s current_rows=%s manual_candidate_rows=%s output=%s yaml_output=%s",
                pure_count,
                current_count,
                diff_count,
                csv_path,
                yaml_path or "",
            )
            return 0
        if args.command == "filter-detail-targets":
            module = load_tool_module("filter_top_yaml_by_manual_apps.py", "netscout_auto_filter_top_yaml_cli")
            top_yaml = module.load_yaml(args.top_yaml)
            manual_yaml = module.load_yaml(args.manual_apps_yaml)
            filtered_yaml, stats = module.filter_top_yaml(
                top_yaml,
                manual_yaml,
                protect_undefined=not args.disable_protected,
            )
            output_path = "" if args.dry_run else str(module.save_yaml(filtered_yaml, args.output))
            logger.info(
                "v0.89 filter-detail-targets completed top_targets=%s manual_apps=%s disabled=%s protected=%s already_disabled=%s output=%s",
                stats["top_targets"],
                stats["manual_apps"],
                stats["disabled"],
                stats["protected"],
                stats["already_disabled"],
                output_path,
            )
            return 0
    except (ConfigError, OSError, RuntimeError, ValueError, V81Error) as exc:
        logger.error("%s", exc)
        return 1
    return 1


def normalize_option_command(argv: list[str]) -> list[str]:
    """플래그형 alias를 argparse 서브커맨드 형태로 변환한다.

    예: `--top --config ...` 입력을 내부적으로 `top --config ...`로 바꾼다.
    """
    option_to_command = {
        "--inventory": "inventory",
        "--top": "top",
        "--detail": "detail",
        "--targets-yaml": "targets-yaml",
        "--run-targets": "run-targets",
        "--detail-targets-yaml": "detail-targets-yaml",
        "--run-detail-targets": "run-detail-targets",
        "--postprocess": "postprocess",
        "--extract-candidates": "extract-candidates",
        "--hostname-lookup": "hostname-lookup",
        "--compare-inventory-apps": "compare-inventory-apps",
        "--filter-detail-targets": "filter-detail-targets",
    }
    normalized = list(argv)
    if normalized and normalized[0] in option_to_command:
        normalized[0] = option_to_command[normalized[0]]
    return normalized


def resolve_env_path(env_path: str) -> str:
    """현재 위치 또는 프로젝트 루트 기준으로 env 파일 경로를 찾는다.

    `.env.local` 내용은 이 함수에서 읽지 않는다. 실제 credential 로드는 공통 config loader가 담당한다.
    """
    path = Path(env_path)
    if path.exists():
        return str(path)
    parent_env = PROJECT_ROOT / env_path
    if parent_env.exists():
        return str(parent_env)
    return env_path


def default_detail_targets_output(top_xlsx: str) -> Path:
    """Top Excel 경로를 기준으로 v0.89 `02_targets/detail_targets.yaml` 기본 경로를 만든다."""
    top_path = Path(top_xlsx)
    if top_path.parent.name == "01_top":
        return top_path.parent.parent / "02_targets" / "detail_targets.yaml"
    return stage_output_dir("output", "02_targets") / "detail_targets.yaml"


def load_tool_module(filename: str, module_name: str) -> Any:
    """Load a script from tools as an internal CLI module."""
    path = PROJECT_ROOT / "tools" / filename
    spec = importlib.util.spec_from_file_location(module_name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load tool module: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def run_postprocess_command(args: argparse.Namespace) -> dict[str, Any]:
    """Run postprocess and optionally continue with hostname lookup."""
    post_module = load_tool_module("postprocess_by_app.py", "netscout_auto_postprocess_by_app_cli")
    result = post_module.run_postprocess_by_app(
        args.date_root,
        targets_yaml=args.targets_yaml,
        include_disabled=args.include_disabled,
        applications=args.applications,
        dry_run=args.dry_run,
    )
    if args.hostname_lookup and not args.dry_run:
        for row in result["results"]:
            yaml_output = row.get("yaml_output")
            if row.get("status") != "success" or not yaml_output:
                continue
            hostname_output = run_hostname_lookup(Path(yaml_output), args.hostname_config, None)
            row["hostname_output"] = str(hostname_output)
    return result


def run_hostname_lookup(input_path: Path, config_path: Path, output_path: Path | None) -> Path:
    """Run hostname lookup from candidate YAML."""
    module = load_tool_module("lookup_hostnames.py", "netscout_auto_lookup_hostnames_cli")
    candidates = module.load_candidates(input_path)
    config = module.load_config(config_path)
    if not candidates:
        raise RuntimeError(f"No valid IP candidates: {input_path}")
    unique_ips = list(dict.fromkeys(row["IP"] for row in candidates))
    lookup_cache = {}
    for index, ip in enumerate(unique_ips, 1):
        logger.info("hostname lookup ip=%s/%s value=%s", index, len(unique_ips), ip)
        lookup_cache[ip] = module.lookup_one_ip(ip, config)
    output = output_path or module.default_output_path(input_path)
    module.write_excel(candidates, lookup_cache, output)
    return output


if __name__ == "__main__":
    raise SystemExit(main())


