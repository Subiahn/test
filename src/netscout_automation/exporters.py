"""정규화된 NETSCOUT 데이터를 XLSX로 저장하는 exporter."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font

from .parsers import APPLICATION_COLUMNS


def export_application_list_xlsx(applications: list[dict[str, Any]], output_path: str | Path) -> Path:
    """Config API Application catalog row를 Application_List sheet에 저장한다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Application_List"
    worksheet.append(APPLICATION_COLUMNS)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for application in applications:
        worksheet.append([application.get(column) for column in APPLICATION_COLUMNS])

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for column_cells in worksheet.columns:
        width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(width + 2, 12), 48)

    workbook.save(path)
    return path
