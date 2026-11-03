"""Inventory Excel 2개를 비교해 수동 등록 Application 후보 CSV/YAML을 만든다.

순정 NETSCOUT에서 뽑은 Inventory의 Application_List를 whitelist로 보고,
앱 등록 후 다시 뽑은 Inventory의 Application_List에서 whitelist와 겹치지 않는
Application만 CSV 또는 YAML로 저장한다. API 호출은 하지 않고 Excel 파일만 읽는다.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path
from typing import Iterable

try:
    from openpyxl import load_workbook
except ModuleNotFoundError:
    # 배포본에서 tools를 직접 실행해도 vendor 패키지를 찾을 수 있게 한다.
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    from openpyxl import load_workbook

try:
    import yaml
except ModuleNotFoundError:
    ROOT = Path(__file__).resolve().parents[1]
    VENDOR = ROOT / "vendor"
    if VENDOR.exists():
        sys.path.insert(0, str(VENDOR))
    import yaml


DEFAULT_SHEET = "Application_List"
DEFAULT_KEY_COLUMN = "ApplicationTypeCode"
DEFAULT_COLUMNS = [
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


class CompareInventoryError(RuntimeError):
    """Inventory 비교 도구에서 복구할 수 없는 오류가 발생했을 때 사용한다."""


def normalize_key(value: object) -> str:
    """Application 비교용 key를 대소문자/공백 차이에 둔감하게 정규화한다."""
    return str(value or "").strip().upper()


def read_sheet_rows(path: Path, sheet_name: str) -> list[dict[str, object]]:
    """Excel의 지정 sheet를 header 기반 dict row 목록으로 읽는다."""
    if not path.exists():
        raise CompareInventoryError(f"Excel file not found: {path}")

    workbook = load_workbook(path, read_only=True, data_only=True)
    if sheet_name not in workbook.sheetnames:
        raise CompareInventoryError(f"Sheet not found: {sheet_name} in {path}")

    worksheet = workbook[sheet_name]
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []

    headers = [str(header or "").strip() for header in rows[0]]
    result: list[dict[str, object]] = []
    for values in rows[1:]:
        row = {headers[index]: value for index, value in enumerate(values) if index < len(headers) and headers[index]}
        if any(value not in (None, "") for value in row.values()):
            result.append(row)
    return result


def build_key_set(rows: Iterable[dict[str, object]], key_column: str) -> set[str]:
    """Application row 목록에서 비교 key set을 만든다."""
    keys: set[str] = set()
    for row in rows:
        key = normalize_key(row.get(key_column))
        if key:
            keys.add(key)
    return keys


def diff_application_rows(
    pure_rows: list[dict[str, object]],
    current_rows: list[dict[str, object]],
    key_column: str = DEFAULT_KEY_COLUMN,
) -> list[dict[str, object]]:
    """current_rows 중 pure_rows에 없는 Application만 반환한다."""
    pure_keys = build_key_set(pure_rows, key_column)
    diff_rows: list[dict[str, object]] = []
    seen: set[str] = set()

    for row in current_rows:
        key = normalize_key(row.get(key_column))
        if not key:
            continue
        if key in pure_keys or key in seen:
            continue
        seen.add(key)
        diff_rows.append(row)

    return diff_rows


def write_csv(rows: list[dict[str, object]], output_path: Path, columns: list[str]) -> Path:
    """비교 결과를 CSV로 저장한다."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({column: row.get(column, "") for column in columns})
    return output_path


def build_yaml_payload(rows: list[dict[str, object]], key_column: str, columns: list[str]) -> dict[str, object]:
    """수동 등록 Application 후보를 사람이 편집하기 쉬운 YAML 구조로 만든다."""
    applications = []
    for row in rows:
        applications.append(
            {
                "enabled": True,
                "key": row.get(key_column),
                "application_short_name": row.get("ApplicationShortName"),
                "application_type_code": row.get("ApplicationTypeCode"),
                "application_long_name": row.get("ApplicationLongName"),
                "application_type": row.get("ApplicationType"),
                "parent_application": row.get("ParentApplication"),
                "application_port": row.get("ApplicationPort"),
                "additional_port": row.get("AdditionalPort"),
                "application_group": row.get("ApplicationGroup"),
                "is_active": row.get("IsActive"),
                "raw": {column: row.get(column) for column in columns},
            }
        )
    return {
        "description": "Pure Inventory에는 없고 Current Inventory에만 있는 수동 등록 Application 후보.",
        "compare_key": key_column,
        "manual_registered_applications": applications,
    }


def write_yaml(rows: list[dict[str, object]], output_path: Path, key_column: str, columns: list[str]) -> Path:
    """비교 결과를 YAML로 저장한다."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = build_yaml_payload(rows, key_column, columns)
    with output_path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(payload, file, allow_unicode=True, sort_keys=False)
    return output_path


def compare_inventory_apps(
    pure_inventory: Path,
    current_inventory: Path,
    output: Path,
    yaml_output: Path | None = None,
    sheet_name: str = DEFAULT_SHEET,
    key_column: str = DEFAULT_KEY_COLUMN,
    columns: list[str] | None = None,
) -> tuple[Path, Path | None, int, int, int]:
    """순정/현재 Inventory Excel을 비교해 현재에만 있는 Application CSV/YAML을 만든다."""
    pure_rows = read_sheet_rows(pure_inventory, sheet_name)
    current_rows = read_sheet_rows(current_inventory, sheet_name)
    diff_rows = diff_application_rows(pure_rows, current_rows, key_column)
    output_columns = columns or DEFAULT_COLUMNS
    csv_path = write_csv(diff_rows, output, output_columns)
    yaml_path = write_yaml(diff_rows, yaml_output, key_column, output_columns) if yaml_output else None
    return csv_path, yaml_path, len(pure_rows), len(current_rows), len(diff_rows)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Compare two NETSCOUT Inventory Excel files and export applications missing from the pure whitelist.",
    )
    parser.add_argument("--pure-inventory", required=True, type=Path, help="순정 NETSCOUT Inventory Excel 경로.")
    parser.add_argument("--current-inventory", required=True, type=Path, help="앱 등록 후 현재 Inventory Excel 경로.")
    parser.add_argument("--output", required=True, type=Path, help="현재 Inventory에만 있는 Application을 저장할 CSV 경로.")
    parser.add_argument("--yaml-output", type=Path, default=None, help="현재 Inventory에만 있는 Application을 저장할 YAML 경로.")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"비교할 sheet 이름. 기본값: {DEFAULT_SHEET}")
    parser.add_argument("--key-column", default=DEFAULT_KEY_COLUMN, help=f"비교 key 컬럼. 기본값: {DEFAULT_KEY_COLUMN}")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        path, yaml_path, pure_count, current_count, diff_count = compare_inventory_apps(
            args.pure_inventory,
            args.current_inventory,
            args.output,
            args.yaml_output,
            args.sheet,
            args.key_column,
        )
    except CompareInventoryError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(
        f"completed pure_rows={pure_count} current_rows={current_count} "
        f"manual_candidate_rows={diff_count} output={path} yaml_output={yaml_path or ''}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
