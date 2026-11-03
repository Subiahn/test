"""저장된 raw Data API JSON probe 응답을 검사하는 helper."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_PROBE_FIELDS = [
    "srcPort",
    "destPort",
    "srcAddress",
    "destAddress",
    "srcToDestOctets",
    "destToSrcOctets",
    "octets",
    "packets",
    "totalTransactions",
]


class RawInspectionError(RuntimeError):
    """raw Data API 응답을 검사할 수 없을 때 발생한다."""


@dataclass(frozen=True)
class FieldInspection:
    field: str
    present_count: int
    non_empty_count: int
    sample_values: list[Any]


@dataclass(frozen=True)
class RawInspectionResult:
    source_path: Path
    row_count: int
    fields: list[FieldInspection]


def inspect_raw_json_fields(raw_json_path: str | Path, fields: list[str] | None = None) -> RawInspectionResult:
    """저장된 raw JSON 응답에 probe 후보 field가 존재하는지 요약한다."""
    path = Path(raw_json_path)
    if not path.exists():
        raise RawInspectionError(f"Raw response file not found: {path}")

    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError as exc:
        raise RawInspectionError(f"Raw response is not valid JSON: {path}") from exc

    rows = _extract_rows(payload)
    target_fields = fields or DEFAULT_PROBE_FIELDS
    inspections = [_inspect_field(rows, field) for field in target_fields]
    return RawInspectionResult(source_path=path, row_count=len(rows), fields=inspections)


def format_raw_inspection(result: RawInspectionResult) -> str:
    lines = [
        f"source_path={result.source_path}",
        f"row_count={result.row_count}",
        "field,present_count,non_empty_count,sample_values",
    ]
    for item in result.fields:
        samples = "|".join(str(value) for value in item.sample_values)
        lines.append(f"{item.field},{item.present_count},{item.non_empty_count},{samples}")
    return "\n".join(lines)


def _extract_rows(payload: Any) -> list[dict[str, Any]]:
    """하나의 API schema로 단정하지 않고 probe 중 확인된 여러 응답 형태를 지원한다."""
    if isinstance(payload, list):
        return [row for row in payload if isinstance(row, dict)]
    if not isinstance(payload, dict):
        raise RawInspectionError("Raw JSON root must be an object or array.")

    for key in ("response", "rows", "data", "outputRows", "outputRow"):
        value = payload.get(key)
        if isinstance(value, list):
            return [row for row in value if isinstance(row, dict)]
        if isinstance(value, dict):
            return [value]
    return [payload]


def _inspect_field(rows: list[dict[str, Any]], field: str) -> FieldInspection:
    present_values = [row.get(field) for row in rows if field in row]
    non_empty_values = [value for value in present_values if value not in (None, "")]
    samples = []
    for value in non_empty_values:
        if value not in samples:
            samples.append(value)
        if len(samples) == 3:
            break
    return FieldInspection(
        field=field,
        present_count=len(present_values),
        non_empty_count=len(non_empty_values),
        sample_values=samples,
    )
