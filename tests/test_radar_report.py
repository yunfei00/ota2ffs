import hashlib

from openpyxl import Workbook

from ota2ffs.matrix_model import PatternMatrix
from ota2ffs.radar_report import add_col_charts, add_row_charts, generate_radar_report


def _matrix(sheet_name="S1", block_name="Theta"):
    return PatternMatrix(
        sheet_name=sheet_name,
        block_name=block_name,
        row_label="Phi Angle",
        col_label="Theta Angle",
        row_angles=[0, 180],
        col_angles=[0, 90, 180],
        values=[[-10, -20, -30], [-40, -50, -60]],
    )


def _add_v2_table(ws, start_row, start_column, block_name, base_value):
    ws.cell(row=start_row, column=start_column, value="Polarization")
    ws.cell(row=start_row, column=start_column + 1, value=block_name)
    ws.cell(row=start_row, column=start_column + 2, value=1800)
    ws.cell(row=start_row + 1, column=start_column, value="Phi\\Theta")
    ws.cell(row=start_row + 1, column=start_column + 1, value=0)
    ws.cell(row=start_row + 1, column=start_column + 2, value=180)
    ws.cell(row=start_row + 2, column=start_column, value=0)
    ws.cell(row=start_row + 2, column=start_column + 1, value=base_value)
    ws.cell(row=start_row + 2, column=start_column + 2, value=base_value - 1)
    ws.cell(row=start_row + 3, column=start_column, value=180)
    ws.cell(row=start_row + 3, column=start_column + 1, value=base_value - 2)
    ws.cell(row=start_row + 3, column=start_column + 2, value=base_value - 3)


def _add_v2_sheet(workbook, title, base_value):
    ws = workbook.create_sheet(title)
    _add_v2_table(ws, 1, 1, "Theta", base_value)
    _add_v2_table(ws, 8, 4, "Phi", base_value - 10)
    _add_v2_table(ws, 16, 2, "Total", base_value - 20)


def test_row_chart_count_equals_row_angle_count():
    workbook = Workbook()
    report_ws = workbook.active
    data_ws = workbook.create_sheet("Data")
    counts = {"single": 0, "compare": 0}

    add_row_charts(report_ws, data_ws, _matrix(), 1, 1, {"row": 1}, counts)

    assert counts["single"] == 2
    assert len(report_ws._charts) == 2


def test_col_chart_count_equals_col_angle_count():
    workbook = Workbook()
    report_ws = workbook.active
    data_ws = workbook.create_sheet("Data")
    counts = {"single": 0, "compare": 0}

    add_col_charts(report_ws, data_ws, _matrix(), 1, 1, {"row": 1}, counts)

    assert counts["single"] == 3
    assert len(report_ws._charts) == 3


def test_multiple_sheets_generate_compare_charts_and_output_xlsx(tmp_path):
    workbook = Workbook()
    workbook.remove(workbook.active)
    _add_v2_sheet(workbook, "S1", -10)
    _add_v2_sheet(workbook, "S2", -20)
    excel_path = tmp_path / "radar_input.xlsx"
    workbook.save(excel_path)

    result = generate_radar_report(excel_path, tmp_path / "out", ["S1", "S2"])

    assert result.output_path.exists()
    assert result.output_path.name == "radar_input_Radar_Report.xlsx"
    assert result.matrix_count == 6
    assert result.single_chart_count == 24
    assert result.compare_chart_count == 12


def test_radar_report_does_not_modify_source_excel(tmp_path):
    workbook = Workbook()
    workbook.remove(workbook.active)
    _add_v2_sheet(workbook, "S1", -10)
    excel_path = tmp_path / "source.xlsx"
    workbook.save(excel_path)
    before_hash = _sha256(excel_path)

    generate_radar_report(excel_path, tmp_path / "out", ["S1"])

    assert _sha256(excel_path) == before_hash


def _sha256(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
