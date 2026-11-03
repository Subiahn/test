"""Application 폴더별 IP/Port 후보 파일을 일괄 생성한다.

이미 생성된 Detail Excel 구조를 기준으로 후처리만 수행한다.
API 호출은 하지 않는다.

예:

output/20261007/03_detail/HTTPS/*.xlsx
  -> output/20261007/04_candidates/HTTPS/ip_port_candidates.xlsx
  -> output/20261007/04_candidates/HTTPS/ip_port_candidates.yaml

output/20261007/03_detail/TCP_OTHER/*.xlsx
  -> output/20261007/04_candidates/TCP_OTHER/ip_port_candidates.xlsx
  -> output/20261007/04_candidates/TCP_OTHER/ip_port_candidates.yaml
"""

from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
TOOLS_DIR = Path(__file__).resolve().parent
VENDOR = ROOT / "vendor"

if VENDOR.exists():
    sys.path.insert(0, str(VENDOR))


def load_extract_module() -> Any:
    """같은 tools 폴더의 extract_ip_port_candidates.py를 파일 경로로 로드한다."""
    path = TOOLS_DIR / "extract_ip_port_candidates.py"
    spec = importlib.util.spec_from_file_location("netscout_auto_extract_ip_port_candidates", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load extract tool: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def app_dirs_from_date_root(date_root: Path) -> list[Path]:
    """날짜 폴더 아래에서 Detail Excel이 들어 있는 Application 폴더만 찾는다."""
    if not date_root.exists():
        raise RuntimeError(f"date root not found: {date_root}")
    if not date_root.is_dir():
        raise RuntimeError(f"date root must be a directory: {date_root}")

    detail_root = date_root / "03_detail"
    search_root = detail_root if detail_root.is_dir() else date_root
    result: list[Path] = []
    for path in sorted(search_root.iterdir(), key=lambda item: item.name.upper()):
        if not path.is_dir():
            continue
        if any(child.suffix.lower() == ".xlsx" and not child.name.startswith("~$") for child in path.rglob("*.xlsx")):
            result.append(path)
    return result


def parse_app_filter(value: str | None) -> set[str]:
    """--applications 값을 대문자 set으로 변환한다."""
    if not value:
        return set()
    return {part.strip().upper() for part in value.split(",") if part.strip()}


def enabled_keys_from_targets(extract_module: Any, targets_yaml: Path | None, include_disabled: bool) -> set[str]:
    """detail targets YAML에서 enabled=true인 Application key만 추출한다."""
    if not targets_yaml:
        return set()
    return extract_module.enabled_application_keys_from_targets(targets_yaml, include_disabled)


def run_postprocess_by_app(
    date_root: Path,
    targets_yaml: Path | None = None,
    include_disabled: bool = False,
    applications: str | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Application 폴더별로 ip_port_candidates.xlsx/yaml을 생성한다."""
    extract_module = load_extract_module()
    app_filter = parse_app_filter(applications)
    target_filter = enabled_keys_from_targets(extract_module, targets_yaml, include_disabled)
    app_dirs = app_dirs_from_date_root(date_root)
    results: list[dict[str, Any]] = []

    for app_dir in app_dirs:
        app_name = app_dir.name
        if app_filter and app_name.upper() not in app_filter:
            results.append({"application": app_name, "status": "skipped", "reason": "application filter"})
            continue
        if target_filter and app_name.upper() not in target_filter:
            results.append({"application": app_name, "status": "skipped", "reason": "not enabled in targets yaml"})
            continue

        candidate_dir = date_root / "04_candidates" / app_name
        output = candidate_dir / "ip_port_candidates.xlsx"
        yaml_output = candidate_dir / "ip_port_candidates.yaml"

        if dry_run:
            results.append(
                {
                    "application": app_name,
                    "status": "planned",
                    "detail_root": str(app_dir),
                    "output": str(output),
                    "yaml_output": str(yaml_output),
                }
            )
            continue

        xlsx_path, yaml_path, stats = extract_module.extract_ip_port_report(
            app_dir,
            output,
            targets_yaml=targets_yaml,
            yaml_output=yaml_output,
            include_disabled=include_disabled,
            applications=app_name,
        )
        results.append(
            {
                "application": app_name,
                "status": "success",
                "detail_root": str(app_dir),
                "output": str(xlsx_path),
                "yaml_output": str(yaml_path or ""),
                "candidate_rows": stats.get("candidate_rows", 0),
                "unique_ip_ports": stats.get("unique_ip_ports", 0),
                "unknown_port_rows": stats.get("unknown_port_rows", 0),
            }
        )

    return {
        "date_root": str(date_root),
        "planned_count": len([row for row in results if row["status"] != "skipped"]),
        "skipped_count": len([row for row in results if row["status"] == "skipped"]),
        "results": results,
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Postprocess NETSCOUT Detail Excel files by application folder.")
    parser.add_argument("--date-root", required=True, type=Path, help="output/YYYYMMDD 날짜 폴더 경로.")
    parser.add_argument("--targets-yaml", type=Path, default=None, help="선택. detail_targets_filtered.yaml 경로.")
    parser.add_argument("--include-disabled", action="store_true", help="선택. targets YAML의 enabled=false 대상도 포함한다.")
    parser.add_argument("--applications", default=None, help="선택. 처리할 Application 목록. 예: HTTPS,TCP_OTHER")
    parser.add_argument("--dry-run", action="store_true", help="선택. 파일을 만들지 않고 처리 예정 대상만 출력한다.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        summary = run_postprocess_by_app(
            args.date_root,
            targets_yaml=args.targets_yaml,
            include_disabled=args.include_disabled,
            applications=args.applications,
            dry_run=args.dry_run,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    for row in summary["results"]:
        print(
            "application={application} status={status} output={output} yaml_output={yaml_output} "
            "candidate_rows={candidate_rows} unique_ip_ports={unique_ip_ports} unknown_port_rows={unknown_port_rows}".format(
                application=row.get("application", ""),
                status=row.get("status", ""),
                output=row.get("output", ""),
                yaml_output=row.get("yaml_output", ""),
                candidate_rows=row.get("candidate_rows", ""),
                unique_ip_ports=row.get("unique_ip_ports", ""),
                unknown_port_rows=row.get("unknown_port_rows", ""),
            )
        )
    print(
        "completed "
        f"date_root={summary['date_root']} "
        f"planned={summary['planned_count']} "
        f"skipped={summary['skipped_count']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
