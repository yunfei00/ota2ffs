from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import RadarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .matrix_model import PatternMatrix
from .matrix_parser import parse_matrices_from_workbook


CHART_WIDTH = 12
CHART_HEIGHT = 8
CHART_COLUMN_STEP = 8
CHART_ROW_STEP = 16
MATRIX_AREA_HEIGHT = 36


@dataclass(slots=True)
class RadarReportResult:
    output_path: Path
    matrix_count: int
    single_chart_count: int
    compare_chart_count: int
    log_lines: list[str] = field(default_factory=list)


def create_radar_report(
    excel_path: str | Path,
    output_dir: str | Path,
    selected_sheets: Iterable[str] | None = None,
) -> Path:
    return generate_radar_report(excel_path, output_dir, selected_sheets).output_path


def generate_radar_report(
    excel_path: str | Path,
    output_dir: str | Path,
    selected_sheets: Iterable[str] | None = None,
) -> RadarReportResult:
    excel_path = Path(excel_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{excel_path.stem}_Radar_Report.xlsx"

    matrices = parse_matrices_from_workbook(excel_path, selected_sheets)
    workbook = Workbook()
    report_ws = workbook.active
    report_ws.title = "Radar_Report"
    data_ws = workbook.create_sheet("Normalized_Data")
    log_ws = workbook.create_sheet("Process_Log")

    chart_counts = {"single": 0, "compare": 0}
    data_cursor = {"row": 1}
    current_row = 1

    for matrix in matrices:
        current_row = add_matrix_area(report_ws, data_ws, matrix, current_row, 1, data_cursor, chart_counts)

    if _has_multiple_sheets(matrices):
        current_row = add_compare_charts(report_ws, data_ws, matrices, current_row, 1, data_cursor, chart_counts)

    log_lines = [
        f"解析矩阵数量: {len(matrices)}",
        f"单图数量: {chart_counts['single']}",
        f"对比图数量: {chart_counts['compare']}",
        f"输出文件: {output_path}",
        "原始 Excel 只读解析，未修改原始文件。",
    ]
    _write_process_log(log_ws, log_lines)

    workbook.save(output_path)
    return RadarReportResult(
        output_path=output_path,
        matrix_count=len(matrices),
        single_chart_count=chart_counts["single"],
        compare_chart_count=chart_counts["compare"],
        log_lines=log_lines,
    )


def add_matrix_area(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrix: PatternMatrix,
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int] | None = None,
    chart_counts: dict[str, int] | None = None,
) -> int:
    if data_cursor is None:
        data_cursor = {"row": 1}
    if chart_counts is None:
        chart_counts = {"single": 0, "compare": 0}

    report_ws.cell(row=start_row, column=start_col, value=f"Sheet: {matrix.sheet_name} / Block: {matrix.block_name}")
    add_row_charts(report_ws, data_ws, matrix, start_row + 1, start_col, data_cursor, chart_counts)
    add_col_charts(report_ws, data_ws, matrix, start_row + CHART_ROW_STEP + 1, start_col, data_cursor, chart_counts)
    return start_row + MATRIX_AREA_HEIGHT


def add_row_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrix: PatternMatrix,
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int],
    chart_counts: dict[str, int],
) -> int:
    normalized = matrix.normalized_values()
    for index, row_angle in enumerate(matrix.row_angles):
        title = f"{matrix.sheet_name}_{matrix.block_name}_Row_{_angle_text(row_angle)}"
        series = [(matrix.sheet_name, normalized[index] if index < len(normalized) else [])]
        chart = _create_chart_from_series(data_ws, data_cursor, title, matrix.col_angles, series)
        _place_chart(report_ws, chart, start_row, start_col + index * CHART_COLUMN_STEP)
        chart_counts["single"] += 1
    return start_row + CHART_ROW_STEP


def add_col_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrix: PatternMatrix,
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int],
    chart_counts: dict[str, int],
) -> int:
    for index, col_angle in enumerate(matrix.col_angles):
        title = f"{matrix.sheet_name}_{matrix.block_name}_Col_{_angle_text(col_angle)}"
        series = [(matrix.sheet_name, matrix.col_values(index))]
        chart = _create_chart_from_series(data_ws, data_cursor, title, matrix.row_angles, series)
        _place_chart(report_ws, chart, start_row, start_col + index * CHART_COLUMN_STEP)
        chart_counts["single"] += 1
    return start_row + CHART_ROW_STEP


def add_compare_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int] | None = None,
    chart_counts: dict[str, int] | None = None,
) -> int:
    if data_cursor is None:
        data_cursor = {"row": 1}
    if chart_counts is None:
        chart_counts = {"single": 0, "compare": 0}

    report_ws.cell(row=start_row, column=start_col, value="Compare Charts")
    _add_compare_row_charts(report_ws, data_ws, matrices, start_row + 1, start_col, data_cursor, chart_counts)
    _add_compare_col_charts(
        report_ws,
        data_ws,
        matrices,
        start_row + CHART_ROW_STEP + 1,
        start_col,
        data_cursor,
        chart_counts,
    )
    return start_row + MATRIX_AREA_HEIGHT


def _add_compare_row_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int],
    chart_counts: dict[str, int],
) -> int:
    count = 0
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        for row_angle in _common_angles([matrix.row_angles for matrix in block_matrices]):
            participating = [matrix for matrix in block_matrices if row_angle in matrix.row_angles]
            if len({matrix.sheet_name for matrix in participating}) < 2:
                continue
            axes = _union_angles([matrix.col_angles for matrix in participating])
            series = [
                (matrix.sheet_name, _row_series_for_axes(matrix, row_angle, axes))
                for matrix in participating
            ]
            title = f"Compare_{block_name}_Row_{_angle_text(row_angle)}"
            chart = _create_chart_from_series(data_ws, data_cursor, title, axes, series)
            _place_chart(report_ws, chart, start_row, start_col + count * CHART_COLUMN_STEP)
            chart_counts["compare"] += 1
            count += 1
    return count


def _add_compare_col_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
    data_cursor: dict[str, int],
    chart_counts: dict[str, int],
) -> int:
    count = 0
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        for col_angle in _common_angles([matrix.col_angles for matrix in block_matrices]):
            participating = [matrix for matrix in block_matrices if col_angle in matrix.col_angles]
            if len({matrix.sheet_name for matrix in participating}) < 2:
                continue
            axes = _union_angles([matrix.row_angles for matrix in participating])
            series = [
                (matrix.sheet_name, _col_series_for_axes(matrix, col_angle, axes))
                for matrix in participating
            ]
            title = f"Compare_{block_name}_Col_{_angle_text(col_angle)}"
            chart = _create_chart_from_series(data_ws, data_cursor, title, axes, series)
            _place_chart(report_ws, chart, start_row, start_col + count * CHART_COLUMN_STEP)
            chart_counts["compare"] += 1
            count += 1
    return count


def _create_chart_from_series(
    data_ws: Worksheet,
    data_cursor: dict[str, int],
    title: str,
    axes: list[float],
    series: list[tuple[str, list[float]]],
) -> RadarChart:
    table_row = data_cursor["row"]
    data_ws.cell(row=table_row, column=1, value=title)
    for index, angle in enumerate(axes, start=2):
        data_ws.cell(row=table_row, column=index, value=_angle_text(angle))

    for row_offset, (series_name, values) in enumerate(series, start=1):
        data_ws.cell(row=table_row + row_offset, column=1, value=series_name)
        value_map = {index: value for index, value in enumerate(values)}
        for index in range(len(axes)):
            data_ws.cell(row=table_row + row_offset, column=index + 2, value=value_map.get(index, 0.0))

    chart = RadarChart()
    chart.type = "standard"
    chart.title = title
    chart.width = CHART_WIDTH
    chart.height = CHART_HEIGHT
    chart.add_data(
        Reference(
            data_ws,
            min_col=1,
            min_row=table_row + 1,
            max_col=len(axes) + 1,
            max_row=table_row + len(series),
        ),
        from_rows=True,
        titles_from_data=True,
    )
    chart.set_categories(Reference(data_ws, min_col=2, min_row=table_row, max_col=len(axes) + 1, max_row=table_row))

    data_cursor["row"] = table_row + len(series) + 3
    return chart


def _place_chart(ws: Worksheet, chart: RadarChart, row: int, column: int) -> None:
    ws.add_chart(chart, f"{get_column_letter(column)}{row}")


def _row_series_for_axes(matrix: PatternMatrix, row_angle: float, axes: list[float]) -> list[float]:
    row_index = matrix.row_angles.index(row_angle)
    row_by_angle = {
        angle: matrix.normalized_values()[row_index][index]
        for index, angle in enumerate(matrix.col_angles)
        if index < len(matrix.normalized_values()[row_index])
    }
    return [row_by_angle.get(angle, 0.0) for angle in axes]


def _col_series_for_axes(matrix: PatternMatrix, col_angle: float, axes: list[float]) -> list[float]:
    col_index = matrix.col_angles.index(col_angle)
    col_by_angle = {
        angle: matrix.col_values(col_index)[index]
        for index, angle in enumerate(matrix.row_angles)
        if index < len(matrix.col_values(col_index))
    }
    return [col_by_angle.get(angle, 0.0) for angle in axes]


def _ordered_block_names(matrices: list[PatternMatrix]) -> list[str]:
    names: list[str] = []
    for matrix in matrices:
        if matrix.block_name not in names:
            names.append(matrix.block_name)
    return names


def _common_angles(angle_lists: list[list[float]]) -> list[float]:
    counts: dict[float, int] = {}
    for angles in angle_lists:
        for angle in set(angles):
            counts[angle] = counts.get(angle, 0) + 1
    return sorted(angle for angle, count in counts.items() if count >= 2)


def _union_angles(angle_lists: list[list[float]]) -> list[float]:
    return sorted({angle for angles in angle_lists for angle in angles})


def _has_multiple_sheets(matrices: list[PatternMatrix]) -> bool:
    return len({matrix.sheet_name for matrix in matrices}) > 1


def _write_process_log(ws: Worksheet, lines: list[str]) -> None:
    ws.cell(row=1, column=1, value="Process Log")
    for index, line in enumerate(lines, start=2):
        ws.cell(row=index, column=1, value=line)


def _angle_text(angle: float) -> str:
    if float(angle).is_integer():
        return str(int(angle))
    return f"{angle:.10f}".rstrip("0").rstrip(".")
