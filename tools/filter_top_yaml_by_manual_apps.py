"""Top 50 Detail targets YAML에서 수동 등록 Application을 enabled=false로 표시한다.

입력은 두 개다.

1. detail-targets-yaml 명령으로 만든 Top 50 YAML
2. compare_inventory_apps.py가 만든 manual_registered_apps.yaml

두 YAML을 ApplicationTypeCode 기준으로 비교하고, 수동 등록 Application과 매칭되는
Top target은 `enabled: false`로 바꾼다. API 호출은 하지 않는다.
"""

from __future__ import annotations

import argparse
import sys
from copy import deepcopy
from pathlib import Path
from typing import Any

try:
    import yaml
except ModuleNotFoundError:
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    import yaml


PROTECTED_KEYWORDS = ("TCP_OTHER", "UDP_OTHER", "IP_OTHER", "OTHER", "UNKNOWN", "IPV6")


class FilterTopYamlError(RuntimeError):
    """Top YAML 필터링 중 복구할 수 없는 오류가 발생했을 때 사용한다."""


def normalize_key(value: object) -> str:
    """Application 비교용 key를 대소문자/공백 차이에 둔감하게 정규화한다."""
    return str(value or "").strip().upper()


def load_yaml(path: Path) -> dict[str, Any]:
    """YAML 파일을 dict로 읽는다."""
    if not path.exists():
        raise FilterTopYamlError(f"YAML file not found: {path}")
    with path.open("r", encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}
    if not isinstance(data, dict):
        raise FilterTopYamlError(f"YAML root must be a mapping: {path}")
    return data


def save_yaml(data: dict[str, Any], path: Path) -> Path:
    """YAML 파일을 저장한다."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(data, file, allow_unicode=True, sort_keys=False)
    return path


def candidate_keys(row: dict[str, Any]) -> set[str]:
    """서로 다른 YAML 구조에서 Application 비교 후보 key를 모은다."""
    keys = set()
    for field in (
        "application_type_code",
        "ApplicationTypeCode",
        "Application Type Code",
        "key",
        "application",
        "Application",
        "application_short_name",
        "ApplicationShortName",
        "Application Name",
    ):
        key = normalize_key(row.get(field))
        if key:
            keys.add(key)
    return keys


def manual_app_keys(manual_yaml: dict[str, Any]) -> set[str]:
    """manual_registered_apps.yaml에서 enabled 상태인 수동 등록 Application key를 추출한다."""
    keys: set[str] = set()
    for item in manual_yaml.get("manual_registered_applications", []):
        if not isinstance(item, dict):
            continue
        if item.get("enabled") is False:
            continue
        keys.update(candidate_keys(item))
    return keys


def is_protected_key(keys: set[str]) -> bool:
    """OTHER/UNKNOWN 계열 등 자동 비활성화 보호 대상인지 확인한다."""
    return any(any(keyword in key for keyword in PROTECTED_KEYWORDS) for key in keys)


def filter_top_yaml(
    top_yaml: dict[str, Any],
    manual_yaml: dict[str, Any],
    protect_undefined: bool = True,
) -> tuple[dict[str, Any], dict[str, int]]:
    """Top target YAML에서 수동 등록 Application target을 enabled=false로 변경한다."""
    output = deepcopy(top_yaml)
    manual_keys = manual_app_keys(manual_yaml)
    targets = output.get("targets", [])
    if not isinstance(targets, list):
        raise FilterTopYamlError("Top YAML field 'targets' must be a list.")

    stats = {
        "top_targets": len(targets),
        "manual_apps": len(manual_keys),
        "disabled": 0,
        "protected": 0,
        "already_disabled": 0,
    }

    for target in targets:
        if not isinstance(target, dict):
            continue
        keys = candidate_keys(target)
        matched = bool(keys & manual_keys)
        target["manual_registered"] = matched
        if not matched:
            continue
        target["manual_registered_keys"] = sorted(keys & manual_keys)
        if target.get("enabled") is False:
            target["disabled_reason"] = target.get("disabled_reason") or "already_disabled"
            stats["already_disabled"] += 1
            continue
        if protect_undefined and is_protected_key(keys):
            target["manual_registered_protected"] = True
            target["disabled_reason"] = "protected_undefined_application"
            stats["protected"] += 1
            continue
        target["enabled"] = False
        target["disabled_reason"] = "manual_registered_application"
        stats["disabled"] += 1

    output["manual_filter"] = {
        "compare_keys": sorted(manual_keys),
        "protect_undefined": protect_undefined,
        "stats": stats,
    }
    return output, stats


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Disable manual registered applications in a Top 50 detail targets YAML.",
    )
    parser.add_argument("--top-yaml", required=True, type=Path, help="detail-targets-yaml로 만든 Top 50 YAML 경로.")
    parser.add_argument("--manual-apps-yaml", required=True, type=Path, help="compare_inventory_apps.py가 만든 수동 등록 App YAML 경로.")
    parser.add_argument("--output", required=True, type=Path, help="필터링된 Top 50 YAML 저장 경로.")
    parser.add_argument("--disable-protected", action="store_true", help="OTHER/UNKNOWN 계열 보호를 끄고 매칭되면 enabled=false 처리한다.")
    parser.add_argument("--dry-run", action="store_true", help="파일 저장 없이 통계만 출력한다.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        top_yaml = load_yaml(args.top_yaml)
        manual_yaml = load_yaml(args.manual_apps_yaml)
        filtered_yaml, stats = filter_top_yaml(top_yaml, manual_yaml, protect_undefined=not args.disable_protected)
        output_path = "" if args.dry_run else str(save_yaml(filtered_yaml, args.output))
    except FilterTopYamlError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(
        "completed "
        f"top_targets={stats['top_targets']} manual_apps={stats['manual_apps']} "
        f"disabled={stats['disabled']} protected={stats['protected']} "
        f"already_disabled={stats['already_disabled']} output={output_path}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
