from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.chart import RadarChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .matrix_model import PatternMatrix
from .matrix_parser import parse_matrices_from_workbook


CHART_WIDTH = 12
CHART_HEIGHT = 8
CHART_COLUMN_STEP = 8
CHART_ROW_STEP = 16
MATRIX_AREA_HEIGHT = 36
DATA_TABLE_GAP_ROWS = 2
DATA_BLOCK_GAP_ROWS = 3
DATA_COMPARE_GAP_COLUMNS = 3

TITLE_FILL = PatternFill("solid", fgColor="1F4E79")
SUBTITLE_FILL = PatternFill("solid", fgColor="D9EAF7")
COMPARE_FILL = PatternFill("solid", fgColor="E2F0D9")
HEADER_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


@dataclass(slots=True)
class RadarReportResult:
    output_path: Path
    matrix_count: int
    single_chart_count: int
    compare_chart_count: int
    log_lines: list[str] = field(default_factory=list)


@dataclass(frozen=True, slots=True)
class DataTableRef:
    start_row: int
    start_col: int
    header_row: int
    first_data_row: int
    row_count: int
    col_count: int

    @property
    def end_row(self) -> int:
        return self.first_data_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.col_count


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
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcOnSave = True
    report_ws = workbook.active
    report_ws.title = "Radar_Report"
    data_ws = workbook.create_sheet("Normalized_Data")
    log_ws = workbook.create_sheet("Process_Log")

    chart_counts = {"single": 0, "compare": 0}
    data_cursor = _prepare_normalized_data(data_ws, matrices, include_compare=_has_multiple_sheets(matrices))
    current_row = 1

    if _has_multiple_sheets(matrices):
        current_row = add_compare_charts(report_ws, data_ws, matrices, current_row, 1, data_cursor, chart_counts)
    else:
        for matrix in matrices:
            current_row = add_matrix_area(report_ws, data_ws, matrix, current_row, 1, data_cursor, chart_counts)

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
    data_cursor: dict[str, object] | None = None,
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
    data_cursor: dict[str, object],
    chart_counts: dict[str, int],
) -> int:
    table = _get_or_write_matrix_table(data_ws, data_cursor, matrix)
    for index, row_angle in enumerate(matrix.row_angles):
        title = f"{matrix.sheet_name}_{matrix.block_name}_Row_{_angle_text(row_angle)}"
        chart = _create_row_chart_from_matrix_table(data_ws, title, matrix.sheet_name, table, index)
        _place_chart(report_ws, chart, start_row, start_col + index * CHART_COLUMN_STEP)
        chart_counts["single"] += 1
    return start_row + CHART_ROW_STEP


def add_col_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrix: PatternMatrix,
    start_row: int,
    start_col: int,
    data_cursor: dict[str, object],
    chart_counts: dict[str, int],
) -> int:
    table = _get_or_write_matrix_table(data_ws, data_cursor, matrix)
    for index, col_angle in enumerate(matrix.col_angles):
        title = f"{matrix.sheet_name}_{matrix.block_name}_Col_{_angle_text(col_angle)}"
        chart = _create_col_chart_from_matrix_table(data_ws, title, matrix.sheet_name, table, index)
        _place_chart(report_ws, chart, start_row, start_col + index * CHART_COLUMN_STEP)
        chart_counts["single"] += 1
    return start_row + CHART_ROW_STEP


def add_compare_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
    data_cursor: dict[str, object] | None = None,
    chart_counts: dict[str, int] | None = None,
) -> int:
    if data_cursor is None:
        data_cursor = {"row": 1}
    if chart_counts is None:
        chart_counts = {"single": 0, "compare": 0}

    report_ws.cell(row=start_row, column=start_col, value="Compare Charts")
    _style_report_title(report_ws, start_row, start_col)

    current_row = start_row + 1
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        report_ws.cell(row=current_row, column=start_col, value=f"Block: {block_name}")
        _style_report_subtitle(report_ws, current_row, start_col)
        _add_compare_row_charts(report_ws, data_ws, block_matrices, current_row + 1, start_col, data_cursor, chart_counts)
        _add_compare_col_charts(
            report_ws,
            data_ws,
            block_matrices,
            current_row + CHART_ROW_STEP + 1,
            start_col,
            data_cursor,
            chart_counts,
        )
        current_row += MATRIX_AREA_HEIGHT
    return current_row


def _add_compare_row_charts(
    report_ws: Worksheet,
    data_ws: Worksheet,
    matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
    data_cursor: dict[str, object],
    chart_counts: dict[str, int],
) -> int:
    count = 0
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        if len({matrix.sheet_name for matrix in block_matrices}) < 2:
            continue
        for row_angle in _common_angles([matrix.row_angles for matrix in block_matrices]):
            title = f"Compare_{block_name}_Row_{_angle_text(row_angle)}"
            table = _get_or_write_compare_table(data_ws, data_cursor, "row", block_name, row_angle, block_matrices)
            chart = _create_chart_from_series_table(data_ws, title, table)
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
    data_cursor: dict[str, object],
    chart_counts: dict[str, int],
) -> int:
    count = 0
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        if len({matrix.sheet_name for matrix in block_matrices}) < 2:
            continue
        for col_angle in _common_angles([matrix.col_angles for matrix in block_matrices]):
            title = f"Compare_{block_name}_Col_{_angle_text(col_angle)}"
            table = _get_or_write_compare_table(data_ws, data_cursor, "col", block_name, col_angle, block_matrices)
            chart = _create_chart_from_series_table(data_ws, title, table)
            _place_chart(report_ws, chart, start_row, start_col + count * CHART_COLUMN_STEP)
            chart_counts["compare"] += 1
            count += 1
    return count


def _prepare_normalized_data(
    ws: Worksheet,
    matrices: list[PatternMatrix],
    include_compare: bool,
) -> dict[str, object]:
    layout: dict[str, object] = {
        "row": 1,
        "matrix_tables": {},
        "compare_tables": {},
    }
    ws.freeze_panes = "B2"

    current_row = 1
    for block_name in _ordered_block_names(matrices):
        block_matrices = [matrix for matrix in matrices if matrix.block_name == block_name]
        max_source_width = max((len(matrix.col_angles) + 1 for matrix in block_matrices), default=1)
        source_col = 1
        compare_col = source_col + max_source_width + DATA_COMPARE_GAP_COLUMNS

        ws.cell(row=current_row, column=source_col, value=f"Block: {block_name}")
        _style_data_section_title(ws, current_row, source_col, max_source_width)
        source_row = current_row + 1
        for matrix in block_matrices:
            table = _write_matrix_table(ws, matrix, source_row, source_col)
            _matrix_tables(layout)[_matrix_key(matrix)] = table
            source_row = table.end_row + DATA_TABLE_GAP_ROWS

        compare_row = current_row + 1
        if include_compare and len({matrix.sheet_name for matrix in block_matrices}) >= 2:
            ws.cell(row=current_row, column=compare_col, value=f"Compare Data: {block_name}")
            _style_data_section_title(ws, current_row, compare_col, max_source_width, compare=True)
            compare_row = _write_compare_tables_for_block(ws, layout, block_matrices, compare_row, compare_col)

        current_row = max(source_row, compare_row) + DATA_BLOCK_GAP_ROWS

    layout["row"] = current_row
    _set_normalized_column_widths(ws)
    return layout


def _write_matrix_table(ws: Worksheet, matrix: PatternMatrix, start_row: int, start_col: int) -> DataTableRef:
    title = f"Sheet: {matrix.sheet_name} / Block: {matrix.block_name}"
    header_row = start_row + 1
    first_data_row = start_row + 2

    ws.cell(row=start_row, column=start_col, value=title)
    _style_data_subtitle(ws, start_row, start_col, len(matrix.col_angles) + 1)
    ws.cell(row=header_row, column=start_col, value=f"{matrix.row_label}\\{matrix.col_label}")
    for offset, angle in enumerate(matrix.col_angles, start=1):
        ws.cell(row=header_row, column=start_col + offset, value=_angle_text(angle))

    normalized = matrix.normalized_values()
    for row_index, row_angle in enumerate(matrix.row_angles):
        row = first_data_row + row_index
        ws.cell(row=row, column=start_col, value=_angle_text(row_angle))
        values = normalized[row_index] if row_index < len(normalized) else []
        for col_index in range(len(matrix.col_angles)):
            value = values[col_index] if col_index < len(values) else 0.0
            ws.cell(row=row, column=start_col + col_index + 1, value=value)

    table = DataTableRef(
        start_row=start_row,
        start_col=start_col,
        header_row=header_row,
        first_data_row=first_data_row,
        row_count=len(matrix.row_angles),
        col_count=len(matrix.col_angles),
    )
    _style_data_table(ws, table)
    return table


def _write_series_table(
    ws: Worksheet,
    title: str,
    axes: list[float],
    series: list[tuple[str, list[float | str]]],
    start_row: int,
    start_col: int,
) -> DataTableRef:
    header_row = start_row + 1
    first_data_row = start_row + 2

    ws.cell(row=start_row, column=start_col, value=title)
    _style_data_subtitle(ws, start_row, start_col, len(axes) + 1, compare=True)
    ws.cell(row=header_row, column=start_col, value="Series")
    for offset, angle in enumerate(axes, start=1):
        ws.cell(row=header_row, column=start_col + offset, value=_angle_text(angle))

    for row_index, (series_name, values) in enumerate(series):
        row = first_data_row + row_index
        ws.cell(row=row, column=start_col, value=series_name)
        for col_index in range(len(axes)):
            value = values[col_index] if col_index < len(values) else 0.0
            ws.cell(row=row, column=start_col + col_index + 1, value=value)

    table = DataTableRef(
        start_row=start_row,
        start_col=start_col,
        header_row=header_row,
        first_data_row=first_data_row,
        row_count=len(series),
        col_count=len(axes),
    )
    _style_data_table(ws, table)
    return table


def _write_compare_tables_for_block(
    ws: Worksheet,
    layout: dict[str, object],
    block_matrices: list[PatternMatrix],
    start_row: int,
    start_col: int,
) -> int:
    current_row = start_row
    for row_angle in _common_angles([matrix.row_angles for matrix in block_matrices]):
        title, axes, series = _compare_row_formula_payload(layout, block_matrices, row_angle)
        table = _write_series_table(ws, title, axes, series, current_row, start_col)
        _compare_tables(layout)[_compare_key("row", block_matrices[0].block_name, row_angle)] = table
        current_row = table.end_row + DATA_TABLE_GAP_ROWS

    for col_angle in _common_angles([matrix.col_angles for matrix in block_matrices]):
        title, axes, series = _compare_col_formula_payload(layout, block_matrices, col_angle)
        table = _write_series_table(ws, title, axes, series, current_row, start_col)
        _compare_tables(layout)[_compare_key("col", block_matrices[0].block_name, col_angle)] = table
        current_row = table.end_row + DATA_TABLE_GAP_ROWS

    return current_row


def _get_or_write_matrix_table(
    data_ws: Worksheet,
    data_cursor: dict[str, object],
    matrix: PatternMatrix,
) -> DataTableRef:
    tables = _matrix_tables(data_cursor)
    key = _matrix_key(matrix)
    if key not in tables:
        start_row = int(data_cursor.get("row", 1))
        table = _write_matrix_table(data_ws, matrix, start_row, 1)
        tables[key] = table
        data_cursor["row"] = table.end_row + DATA_TABLE_GAP_ROWS
    return tables[key]


def _get_or_write_compare_table(
    data_ws: Worksheet,
    data_cursor: dict[str, object],
    kind: str,
    block_name: str,
    angle: float,
    block_matrices: list[PatternMatrix],
) -> DataTableRef:
    tables = _compare_tables(data_cursor)
    key = _compare_key(kind, block_name, angle)
    if key not in tables:
        if _has_source_matrix_tables(data_cursor, block_matrices):
            title, axes, series = (
                _compare_row_formula_payload(data_cursor, block_matrices, angle)
                if kind == "row"
                else _compare_col_formula_payload(data_cursor, block_matrices, angle)
            )
        else:
            title, axes, series = (
                _compare_row_payload(block_matrices, angle)
                if kind == "row"
                else _compare_col_payload(block_matrices, angle)
            )
        start_row = int(data_cursor.get("row", 1))
        table = _write_series_table(data_ws, title, axes, series, start_row, 1)
        tables[key] = table
        data_cursor["row"] = table.end_row + DATA_TABLE_GAP_ROWS
    return tables[key]


def _create_row_chart_from_matrix_table(
    data_ws: Worksheet,
    title: str,
    series_name: str,
    table: DataTableRef,
    row_index: int,
) -> RadarChart:
    data_row = table.first_data_row + row_index
    chart = _new_radar_chart(title)
    chart.add_data(
        Reference(
            data_ws,
            min_col=table.start_col,
            min_row=data_row,
            max_col=table.end_col,
            max_row=data_row,
        ),
        from_rows=True,
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(data_ws, min_col=table.start_col + 1, min_row=table.header_row, max_col=table.end_col)
    )
    _set_series_titles(chart, [series_name])
    return chart


def _create_col_chart_from_matrix_table(
    data_ws: Worksheet,
    title: str,
    series_name: str,
    table: DataTableRef,
    col_index: int,
) -> RadarChart:
    value_col = table.start_col + col_index + 1
    chart = _new_radar_chart(title)
    chart.add_data(
        Reference(
            data_ws,
            min_col=value_col,
            min_row=table.header_row,
            max_col=value_col,
            max_row=table.end_row,
        ),
        titles_from_data=True,
    )
    chart.set_categories(Reference(data_ws, min_col=table.start_col, min_row=table.first_data_row, max_row=table.end_row))
    _set_series_titles(chart, [series_name])
    return chart


def _create_chart_from_series_table(data_ws: Worksheet, title: str, table: DataTableRef) -> RadarChart:
    chart = _new_radar_chart(title)
    chart.add_data(
        Reference(
            data_ws,
            min_col=table.start_col,
            min_row=table.first_data_row,
            max_col=table.end_col,
            max_row=table.end_row,
        ),
        from_rows=True,
        titles_from_data=True,
    )
    chart.set_categories(
        Reference(data_ws, min_col=table.start_col + 1, min_row=table.header_row, max_col=table.end_col)
    )
    return chart


def _new_radar_chart(title: str) -> RadarChart:
    chart = RadarChart()
    chart.type = "standard"
    chart.title = title
    chart.width = CHART_WIDTH
    chart.height = CHART_HEIGHT
    return chart


def _set_series_titles(chart: RadarChart, names: list[str]) -> None:
    for series, name in zip(chart.series, names):
        series.tx = SeriesLabel(v=name)


def _place_chart(ws: Worksheet, chart: RadarChart, row: int, column: int) -> None:
    ws.add_chart(chart, f"{get_column_letter(column)}{row}")


def _row_series_for_axes(matrix: PatternMatrix, row_angle: float, axes: list[float]) -> list[float]:
    if row_angle not in matrix.row_angles:
        return [0.0 for _ in axes]
    row_index = matrix.row_angles.index(row_angle)
    normalized = matrix.normalized_values()
    row_values = normalized[row_index] if row_index < len(normalized) else []
    row_by_angle = {
        angle: row_values[index]
        for index, angle in enumerate(matrix.col_angles)
        if index < len(row_values)
    }
    return [row_by_angle.get(angle, 0.0) for angle in axes]


def _col_series_for_axes(matrix: PatternMatrix, col_angle: float, axes: list[float]) -> list[float]:
    if col_angle not in matrix.col_angles:
        return [0.0 for _ in axes]
    col_index = matrix.col_angles.index(col_angle)
    col_values = matrix.col_values(col_index)
    col_by_angle = {
        angle: col_values[index]
        for index, angle in enumerate(matrix.row_angles)
        if index < len(col_values)
    }
    return [col_by_angle.get(angle, 0.0) for angle in axes]


def _compare_row_payload(
    block_matrices: list[PatternMatrix],
    row_angle: float,
) -> tuple[str, list[float], list[tuple[str, list[float]]]]:
    block_name = block_matrices[0].block_name
    axes = _union_angles([matrix.col_angles for matrix in block_matrices])
    series = [(matrix.sheet_name, _row_series_for_axes(matrix, row_angle, axes)) for matrix in block_matrices]
    return f"Compare_{block_name}_Row_{_angle_text(row_angle)}", axes, series


def _compare_col_payload(
    block_matrices: list[PatternMatrix],
    col_angle: float,
) -> tuple[str, list[float], list[tuple[str, list[float]]]]:
    block_name = block_matrices[0].block_name
    axes = _union_angles([matrix.row_angles for matrix in block_matrices])
    series = [(matrix.sheet_name, _col_series_for_axes(matrix, col_angle, axes)) for matrix in block_matrices]
    return f"Compare_{block_name}_Col_{_angle_text(col_angle)}", axes, series


def _compare_row_formula_payload(
    layout: dict[str, object],
    block_matrices: list[PatternMatrix],
    row_angle: float,
) -> tuple[str, list[float], list[tuple[str, list[float | str]]]]:
    block_name = block_matrices[0].block_name
    axes = _union_angles([matrix.col_angles for matrix in block_matrices])
    series = [
        (matrix.sheet_name, [_row_formula_or_zero(layout, matrix, row_angle, col_angle) for col_angle in axes])
        for matrix in block_matrices
    ]
    return f"Compare_{block_name}_Row_{_angle_text(row_angle)}", axes, series


def _compare_col_formula_payload(
    layout: dict[str, object],
    block_matrices: list[PatternMatrix],
    col_angle: float,
) -> tuple[str, list[float], list[tuple[str, list[float | str]]]]:
    block_name = block_matrices[0].block_name
    axes = _union_angles([matrix.row_angles for matrix in block_matrices])
    series = [
        (matrix.sheet_name, [_col_formula_or_zero(layout, matrix, col_angle, row_angle) for row_angle in axes])
        for matrix in block_matrices
    ]
    return f"Compare_{block_name}_Col_{_angle_text(col_angle)}", axes, series


def _row_formula_or_zero(
    layout: dict[str, object],
    matrix: PatternMatrix,
    row_angle: float,
    col_angle: float,
) -> float | str:
    if row_angle not in matrix.row_angles or col_angle not in matrix.col_angles:
        return 0.0
    table = _matrix_tables(layout).get(_matrix_key(matrix))
    if table is None:
        return _row_series_for_axes(matrix, row_angle, [col_angle])[0]
    row = table.first_data_row + matrix.row_angles.index(row_angle)
    col = table.start_col + matrix.col_angles.index(col_angle) + 1
    return _cell_formula(row, col)


def _col_formula_or_zero(
    layout: dict[str, object],
    matrix: PatternMatrix,
    col_angle: float,
    row_angle: float,
) -> float | str:
    if col_angle not in matrix.col_angles or row_angle not in matrix.row_angles:
        return 0.0
    table = _matrix_tables(layout).get(_matrix_key(matrix))
    if table is None:
        return _col_series_for_axes(matrix, col_angle, [row_angle])[0]
    row = table.first_data_row + matrix.row_angles.index(row_angle)
    col = table.start_col + matrix.col_angles.index(col_angle) + 1
    return _cell_formula(row, col)


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


def _matrix_key(matrix: PatternMatrix) -> tuple[int, str, str]:
    return (id(matrix), matrix.sheet_name, matrix.block_name)


def _compare_key(kind: str, block_name: str, angle: float) -> tuple[str, str, float]:
    return kind, block_name, angle


def _matrix_tables(layout: dict[str, object]) -> dict[tuple[int, str, str], DataTableRef]:
    return layout.setdefault("matrix_tables", {})  # type: ignore[return-value]


def _compare_tables(layout: dict[str, object]) -> dict[tuple[str, str, float], DataTableRef]:
    return layout.setdefault("compare_tables", {})  # type: ignore[return-value]


def _has_source_matrix_tables(layout: dict[str, object], matrices: list[PatternMatrix]) -> bool:
    tables = _matrix_tables(layout)
    return all(_matrix_key(matrix) in tables for matrix in matrices)


def _cell_formula(row: int, col: int) -> str:
    return f"=${get_column_letter(col)}${row}"


def _style_report_title(ws: Worksheet, row: int, col: int) -> None:
    cell = ws.cell(row=row, column=col)
    cell.font = Font(bold=True, size=14)


def _style_report_subtitle(ws: Worksheet, row: int, col: int) -> None:
    cell = ws.cell(row=row, column=col)
    cell.font = Font(bold=True, size=12)


def _style_data_section_title(
    ws: Worksheet,
    row: int,
    col: int,
    width: int,
    compare: bool = False,
) -> None:
    for column in range(col, col + max(width, 1)):
        cell = ws.cell(row=row, column=column)
        cell.fill = COMPARE_FILL if compare else TITLE_FILL
        cell.font = Font(bold=True, color="000000" if compare else "FFFFFF")
        cell.alignment = Alignment(horizontal="center")


def _style_data_subtitle(
    ws: Worksheet,
    row: int,
    col: int,
    width: int,
    compare: bool = False,
) -> None:
    for column in range(col, col + max(width, 1)):
        cell = ws.cell(row=row, column=column)
        cell.fill = COMPARE_FILL if compare else SUBTITLE_FILL
        cell.font = Font(bold=True)


def _style_data_table(ws: Worksheet, table: DataTableRef) -> None:
    for row in range(table.header_row, table.end_row + 1):
        for col in range(table.start_col, table.end_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if row == table.header_row:
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)


def _set_normalized_column_widths(ws: Worksheet) -> None:
    for column in range(1, ws.max_column + 1):
        letter = get_column_letter(column)
        ws.column_dimensions[letter].width = 12
    ws.column_dimensions["A"].width = 22
